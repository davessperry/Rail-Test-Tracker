#!/usr/bin/env python3
"""Builds DailyReport.oxt (a LibreOffice extension) from DailyReport.bas.

Double-clicking the .oxt installs the macro library and adds a "Daily Report" menu to
Calc (Fill Remaining Miles / Fill Miles Tested). Run:  python3 build_oxt.py
"""
import zipfile, html, os
here = os.path.dirname(os.path.abspath(__file__))
code = open(os.path.join(here, 'DailyReport.bas'), encoding='utf-8').read().replace('\r\n', '\n')

module = ('<?xml version="1.0" encoding="UTF-8"?>\n'
 '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
 '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">'
 + html.escape(code, quote=False) + '</script:module>\n')
script_xlb = ('<?xml version="1.0" encoding="UTF-8"?>\n'
 '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
 '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="DailyReport" library:readonly="false" library:passwordprotected="false">\n'
 ' <library:element library:name="Module1"/>\n</library:library>\n')
dialog_xlb = ('<?xml version="1.0" encoding="UTF-8"?>\n'
 '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
 '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="DailyReport" library:readonly="false" library:passwordprotected="false"/>\n')
manifest = ('<?xml version="1.0" encoding="UTF-8"?>\n'
 '<manifest:manifest xmlns:manifest="http://openoffice.org/2001/manifest">\n'
 ' <manifest:file-entry manifest:media-type="application/vnd.sun.star.basic-library" manifest:full-path="DailyReport/"/>\n'
 ' <manifest:file-entry manifest:media-type="application/vnd.sun.star.configuration-data" manifest:full-path="Addons.xcu"/>\n'
 '</manifest:manifest>\n')
description = ('<?xml version="1.0" encoding="UTF-8"?>\n'
 '<description xmlns="http://openoffice.org/extensions/description/2006" xmlns:xlink="http://www.w3.org/1999/xlink">\n'
 ' <identifier value="org.railtesttracker.dailyreport"/>\n'
 ' <version value="1.0.0"/>\n'
 ' <display-name><name lang="en">Rail Test Tracker - Daily Report</name></display-name>\n'
 '</description>\n')

def item(n, title, macro):
    return ('     <node oor:name="%s" oor:op="replace">\n'
     '      <prop oor:name="URL" oor:type="xs:string"><value>vnd.sun.star.script:DailyReport.Module1.%s?language=Basic&amp;location=application</value></prop>\n'
     '      <prop oor:name="Title" oor:type="xs:string"><value xml:lang="en-US">%s</value></prop>\n'
     '      <prop oor:name="Target" oor:type="xs:string"><value>_self</value></prop>\n'
     '     </node>\n') % (n, macro, title)
addons = ('<?xml version="1.0" encoding="UTF-8"?>\n'
 '<oor:component-data xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" oor:name="Addons" oor:package="org.openoffice.Office">\n'
 ' <node oor:name="AddonUI">\n'
 '  <node oor:name="OfficeMenuBar">\n'
 '   <node oor:name="org.railtesttracker.dailyreport" oor:op="replace">\n'
 '    <prop oor:name="Title" oor:type="xs:string"><value xml:lang="en-US">Daily Report</value></prop>\n'
 '    <prop oor:name="Target" oor:type="xs:string"><value>_self</value></prop>\n'
 '    <prop oor:name="Context" oor:type="xs:string"><value>com.sun.star.sheet.SpreadsheetDocument</value></prop>\n'
 '    <node oor:name="Submenu">\n'
 + item('m1', 'Fill Remaining Miles', 'FillRemaining') + item('m2', 'Fill Miles Tested Today', 'FillTested') +
 '    </node>\n'
 '   </node>\n'
 '  </node>\n'
 ' </node>\n'
 '</oor:component-data>\n')

out = os.path.join(here, 'DailyReport.oxt')
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('META-INF/manifest.xml', manifest)
    z.writestr('description.xml', description)
    z.writestr('Addons.xcu', addons)
    z.writestr('DailyReport/script.xlb', script_xlb)
    z.writestr('DailyReport/dialog.xlb', dialog_xlb)
    z.writestr('DailyReport/Module1.xba', module)
print('wrote', out, os.path.getsize(out), 'bytes')
