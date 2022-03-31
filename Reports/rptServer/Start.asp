<%@ LANGUAGE="VBSCRIPT" %>
<%
' Copyright © 1997-2002 Crystal Decisions, Inc.
' Modified for the purposes of ePortfolio Lite

Option Explicit

On Error Resume Next

'Get the report name and put it into the session
Dim theReportName
theReportName = Request.Form("ReportName")
if theReportName = "" then theReportName = Request.QueryString("ReportName")
Session("ReportName") = theReportName

' INSTANTIATE THE CRYSTAL REPORTS VIEWER
'
'The are three Crystal Reports Smart Viewers:
'
'1.  ActiveX Smart Viewer
'2.  Java Viewer using the plug-in JVM
'3.  Netscape Plug-in Viewer
'
'The Viewer that you use will based on the browser's display capablities.
'For Example, you would not want to instantiate the Java viewer if the browser
'did not support Java applets.  For purposes on this demo, we have chosen to
'define a viewer.  You can through code determine the support capabilities of
'the requesting browser.  However that functionality is inherent in the Crystal
'Reports automation server and is beyond the scope of this demonstration app.
'
'We have chosen to leverage the server side include functionality of ASP
'for simplicity sake.  So you can use the *Viewer.asp files to instantiate
'the smart viewer that you wish to send to the browser.  Simply replace the line
'below with the Viewer asp file you wish to use.
'
'The choices are ViewerActiveX.asp, JavaViewer.asp, NetscapePluginViewer.asp
'and JavaPluginViewer.asp.
'Note that to use this include you must have the appropriate .asp file in the 
'same virtual directory as the main ASP page.
'
Dim viewer
viewer = Request.Form("Viewer")
if viewer = "" then viewer = Request.QueryString("Viewer")
If viewer = "0" Then
%>
<!-- #include file="ActiveXViewer.asp" -->
<%
ElseIf viewer = "3" Then
%>
<!-- #include file="JavaPluginViewer.asp" -->
<%
End If
%>
