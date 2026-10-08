REM Rail Test Tracker -> Daily Report (LibreOffice Calc macro)
REM In the app, tap "Copy Remaining Miles" or "Copy Miles Tested Today", switch to the
REM spreadsheet and run FillRemaining or FillTested. The macro reads what is on the
REM clipboard, finds the label cells in column B, clears the old data to their right
REM (columns C and D), writes the new data, and adds rows when it does not fit.
REM Nothing is deleted from column B or from any other column.

Option Explicit

REM ---- settings you can change if your sheet changes ----
Const LABEL_COL = 1        ' column B (A=0, B=1, ...) holds the labels
Const DATA_COL = 2         ' data goes in column C and D
Const MAX_BLOCK_ROWS = 80  ' safety stop: never clear more rows than this
Const TOTAL_LABEL = "total miles left"
Const REMAINING_LABEL = "range of miles left"
Const TESTED_LABEL = "range of miles tested"

Sub FillRemaining
  Show ApplyText("remaining", GetClipboardText())
End Sub

Sub FillTested
  Show ApplyText("tested", GetClipboardText())
End Sub

Sub Show(sMsg As String)
  If Left(sMsg, 5) = "ERROR" Or Left(sMsg, 4) = "NOTE" Then
    MsgBox sMsg, 48, "Daily report"
  Else
    MsgBox sMsg, 64, "Daily report"
  End If
End Sub

Function GetClipboardText() As String
  Dim oClip As Object, oContents As Object, aFlavors As Variant, i As Long
  GetClipboardText = ""
  oClip = CreateUnoService("com.sun.star.datatransfer.clipboard.SystemClipboard")
  oContents = oClip.getContents()
  aFlavors = oContents.getTransferDataFlavors()
  For i = LBound(aFlavors) To UBound(aFlavors)
    If aFlavors(i).MimeType = "text/plain;charset=utf-16" Then
      GetClipboardText = oContents.getTransferData(aFlavors(i))
      Exit Function
    End If
  Next i
  ' fallback: any other plain-text flavor
  For i = LBound(aFlavors) To UBound(aFlavors)
    If Left(aFlavors(i).MimeType, 10) = "text/plain" Then
      GetClipboardText = oContents.getTransferData(aFlavors(i))
      Exit Function
    End If
  Next i
End Function

REM Works out the pieces from the copied text and writes them into the sheet.
REM Returns a short message (shown to the user by the macros above).
Function ApplyText(sKind As String, sText As String) As String
  Dim aLines() As String, aSum() As String, aBody() As String
  Dim nSum As Long, nBody As Long, i As Long, k As Long
  Dim sLine As String, sMsg As String, bHasBracket As Boolean

  If Len(Trim(sText)) = 0 Then
    ApplyText = "NOTE: The clipboard is empty. Tap a Copy button in the app first."
    Exit Function
  End If
  sText = Replace(sText, Chr(13) & Chr(10), Chr(10))
  sText = Replace(sText, Chr(13), Chr(10))
  aLines = Split(sText, Chr(10))
  ' drop trailing empty lines
  k = UBound(aLines)
  Do While k >= 0
    If Len(Replace(aLines(k), Chr(9), "")) > 0 Then Exit Do
    k = k - 1
  Loop
  If k < 0 Then
    ApplyText = "NOTE: The clipboard has nothing to write."
    Exit Function
  End If

  ReDim aSum(k)
  ReDim aBody(k)
  nSum = 0: nBody = 0
  i = 0
  ' leading "Name: total plus ..." lines are the total lines
  Do While i <= k
    If IsTotalLine(aLines(i)) Then
      aSum(nSum) = aLines(i): nSum = nSum + 1: i = i + 1
    Else
      Exit Do
    End If
  Loop
  ' skip the blank spacer row(s) between the totals and the first subdivision
  Do While i <= k
    If Len(Replace(aLines(i), Chr(9), "")) = 0 And nSum > 0 And nBody = 0 Then
      i = i + 1
    Else
      Exit Do
    End If
  Loop
  For i = i To k
    aBody(nBody) = aLines(i): nBody = nBody + 1
    If InStr(aLines(i), "[") > 0 Then bHasBracket = True
  Next i

  ' guard against sending the wrong kind of data to the wrong block
  If sKind = "remaining" And nSum = 0 And Not bHasBracket Then
    ApplyText = "NOTE: That does not look like Remaining Miles data. Tap 'Copy Remaining Miles' in the app first."
    Exit Function
  End If
  If sKind = "tested" And (nSum > 0 Or bHasBracket) Then
    ApplyText = "NOTE: That looks like Remaining Miles data. Tap 'Copy Miles Tested Today' in the app first."
    Exit Function
  End If

  If sKind = "remaining" Then
    sMsg = WriteBlock(TOTAL_LABEL, REMAINING_LABEL, aSum, nSum, True)
    If Left(sMsg, 5) <> "ERROR" Then
      sMsg = sMsg & Chr(10) & WriteBlock(REMAINING_LABEL, "", aBody, nBody, False)
    End If
  Else
    sMsg = WriteBlock(TESTED_LABEL, "", aBody, nBody, False)
  End If
  ApplyText = sMsg
End Function

REM A total line looks like "Cascade: 94.3 plus sidings" with nothing in the 2nd column.
Function IsTotalLine(s As String) As Boolean
  Dim aCells() As String, sFirst As String, p As Long, sRest As String
  IsTotalLine = False
  aCells = Split(s, Chr(9))
  If UBound(aCells) >= 1 Then
    If Len(Trim(aCells(1))) > 0 Then Exit Function
  End If
  sFirst = Trim(aCells(0))
  p = InStr(sFirst, ": ")
  If p < 2 Then Exit Function
  sRest = Trim(Mid(sFirst, p + 2))
  If Len(sRest) = 0 Then Exit Function
  If InStr("0123456789", Left(sRest, 1)) = 0 Then Exit Function
  IsTotalLine = True
End Function

REM A subdivision heading: ALL CAPS in the first cell and nothing in the second.
Function IsTitleLine(s As String) As Boolean
  Dim aCells() As String, sFirst As String
  IsTitleLine = False
  aCells = Split(s, Chr(9))
  If UBound(aCells) >= 1 Then
    If Len(Trim(aCells(1))) > 0 Then Exit Function
  End If
  sFirst = Trim(aCells(0))
  If Len(sFirst) = 0 Then Exit Function
  If InStr(sFirst, "[") > 0 Then Exit Function
  If UCase(sFirst) = sFirst And LCase(sFirst) <> sFirst Then IsTitleLine = True
End Function

Function NormText(s As String) As String
  Dim t As String
  t = LCase(Trim(s))
  t = Replace(t, ":", " ")
  t = Replace(t, Chr(160), " ")
  Do While InStr(t, "  ") > 0
    t = Replace(t, "  ", " ")
  Loop
  NormText = Trim(t)
End Function

REM Row index (0-based) of the first label containing sNeedle in column B, or -1.
Function FindLabelRow(oSheet As Object, sNeedle As String, nFrom As Long) As Long
  Dim oCur As Object, nLast As Long, r As Long
  oCur = oSheet.createCursor()
  oCur.gotoEndOfUsedArea(False)
  nLast = oCur.RangeAddress.EndRow
  FindLabelRow = -1
  For r = nFrom To nLast
    If InStr(NormText(oSheet.getCellByPosition(LABEL_COL, r).String), sNeedle) > 0 Then
      FindLabelRow = r
      Exit Function
    End If
  Next r
End Function

Function LastUsedRow(oSheet As Object) As Long
  Dim oCur As Object
  oCur = oSheet.createCursor()
  oCur.gotoEndOfUsedArea(False)
  LastUsedRow = oCur.RangeAddress.EndRow
End Function

REM Fills one block. sStart = label text the block starts at; sEnd = label text where it
REM ends ("" = up to the next non-empty cell in column B, or the end of the sheet).
Function WriteBlock(sStart As String, sEnd As String, aRows() As String, nRows As Long, bMergeAll As Boolean) As String
  Dim oSheet As Object, oSheets As Object, nStart As Long, nEnd As Long, nLast As Long
  Dim s As Long, r As Long, nHeight As Long, nExtra As Long, aCells() As String
  Dim oRange As Object, oCell As Object

  oSheets = ThisComponent.Sheets
  nStart = -1
  For s = 0 To oSheets.Count - 1
    r = FindLabelRow(oSheets.getByIndex(s), sStart, 0)
    If r >= 0 Then
      oSheet = oSheets.getByIndex(s): nStart = r: Exit For
    End If
  Next s
  If nStart < 0 Then
    WriteBlock = "ERROR: could not find a label containing '" & sStart & "' in column B."
    Exit Function
  End If

  nLast = LastUsedRow(oSheet)
  nEnd = -1
  If Len(sEnd) > 0 Then
    r = FindLabelRow(oSheet, sEnd, nStart + 1)
    If r > nStart Then nEnd = r
  End If
  If nEnd < 0 Then
    For r = nStart + 1 To nLast
      If Len(Trim(oSheet.getCellByPosition(LABEL_COL, r).String)) > 0 Then nEnd = r: Exit For
    Next r
  End If
  If nEnd < 0 Then nEnd = nLast + 1
  nHeight = nEnd - nStart
  If nHeight > MAX_BLOCK_ROWS Then
    WriteBlock = "ERROR: the block under '" & sStart & "' looks too tall (" & nHeight & " rows). Nothing was changed."
    Exit Function
  End If

  ' add whole rows if the data does not fit (new rows copy the look of the row above)
  If nRows > nHeight Then
    nExtra = nRows - nHeight
    oSheet.Rows.insertByIndex(nEnd, nExtra)
    nEnd = nEnd + nExtra
    nHeight = nHeight + nExtra
  End If
  If nHeight < 1 Then nHeight = 1

  ' clear the old data (columns C and D only)
  oRange = oSheet.getCellRangeByPosition(DATA_COL, nStart, DATA_COL + 1, nStart + nHeight - 1)
  oRange.merge(False)
  oRange.clearContents(1 + 2 + 4 + 16)   ' values, dates, strings, formulas
  oRange.CharWeight = 100                ' normal
  oRange.CharUnderline = 0               ' none

  For r = 0 To nRows - 1
    aCells = Split(aRows(r), Chr(9))
    oCell = oSheet.getCellByPosition(DATA_COL, nStart + r)
    If UBound(aCells) >= 0 Then oCell.String = Trim(aCells(0))
    If UBound(aCells) >= 1 Then oSheet.getCellByPosition(DATA_COL + 1, nStart + r).String = Trim(aCells(1))
    If bMergeAll Then
      oSheet.getCellRangeByPosition(DATA_COL, nStart + r, DATA_COL + 1, nStart + r).merge(True)
    ElseIf IsTitleLine(aRows(r)) Then
      oCell.CharWeight = 150             ' bold
      oCell.CharUnderline = 1            ' single
    End If
  Next r
  WriteBlock = "Wrote " & nRows & " rows next to '" & sStart & "' (row " & (nStart + 1) & ")."
End Function
