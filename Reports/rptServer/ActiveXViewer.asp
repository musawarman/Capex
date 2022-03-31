<%
' Copyright © 1997-2002 Crystal Decisions, Inc.
'
'This file contains the HTML code to instantiate the Crystal Reports ActiveX Viewer.
'                                                                     
'You will notice that the Report Name parameter references the rptserver.asp file.
'This is because the report pages are actually created by rptserver.asp.
'Rptserver.asp accesses session("oClientDoc") to create the report pages 
'that will be rendered by the ActiveX Viewer.
'
%>
<HTML>
<HEAD>
<TITLE>Crystal Reports ActiveX Viewer</TITLE>
</HEAD>
<BODY BGCOLOR=C6C6C6 onunload="CallDestroy()" topmargin=0 leftmargin=0>

<OBJECT ID="CRViewer"
	CLASSID="CLSID:2DEF4530-8CE6-41c9-84B6-A54536C90213"
	WIDTH=100% HEIGHT=99%
	CODEBASE="/viewer9/activeXViewer/activexviewer.cab#Version=9,2,0,442" VIEWASTEXT>
<PARAM NAME="EnableRefreshButton" VALUE=1>
<PARAM NAME="EnableGroupTree" VALUE=1>
<PARAM NAME="DisplayGroupTree" VALUE=1>
<PARAM NAME="EnablePrintButton" VALUE=1>
<PARAM NAME="EnableExportButton" VALUE=1>
<PARAM NAME="EnableDrillDown" VALUE=1>
<PARAM NAME="EnableSearchControl" VALUE=1>
<PARAM NAME="EnableAnimationControl" VALUE=1>
<PARAM NAME="EnableZoomControl" VALUE=1>
<PARAM NAME="EnableSelectExpertButton" VALUE=1>
</OBJECT>

<SCRIPT LANGUAGE="VBScript">
<!--
Sub Window_Onload
	On Error Resume Next
	Dim webBroker
	Set webBroker = CreateObject("WebReportBroker9.WebReportBroker")
	If ScriptEngineMajorVersion < 2 Then
		window.alert "IE 3.02 users need to get the latest version of VBScript or install IE 4.01 SP1 or newer. Users of Windows 95 additionally need DCOM95.  These files are available at Microsoft's web site."
	else
		Dim webSource
		Set webSource = CreateObject("WebReportSource9.WebReportSource")
		webSource.ReportSource = webBroker
		webSource.URL = "rptserver.asp"
		webSource.PromptOnRefresh = True
		CRViewer.ReportSource = webSource
	End If
	CRViewer.ViewReport
End Sub
-->
</SCRIPT>

<script language="javascript">
function CallDestroy()
{
	window.open("Cleanup.asp");
}
</script>

<OBJECT ID="ReportSource"
	CLASSID="CLSID:934CC260-C5AA-43C4-A657-7B70C5B3DAE1"
	HEIGHT=1% WIDTH=1%
    CODEBASE="/viewer9/activeXViewer/activexviewer.cab#Version=9,2,0,442">
</OBJECT>
<OBJECT ID="ViewHelp"
	CLASSID="CLSID:4B5C9C28-3806-47b5-89A9-93063323160F"
	HEIGHT=1% WIDTH=1%
    CODEBASE="/viewer9/activeXViewer/activexviewer.cab#Version=9,2,0,442">
</OBJECT>
<div>
<!-- This empty div prevents IE from showing a bunch of empty space for the controls above. -->
</div>

</BODY>
</HTML>
