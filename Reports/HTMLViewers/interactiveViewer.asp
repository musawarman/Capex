<%@CodePage=65001%>
<% 
Option Explicit
Response.ContentType = "text/html; charset=utf-8"

Dim objectFactory
Set objectFactory = CreateObject("CrystalReports.ObjectFactory.2")

Response.ExpiresAbsolute = Now() - 1
	
Dim viewer
Set viewer = objectFactory.CreateObject("CrystalReports.CrystalReportInteractiveViewer")  
viewer.Name = "page"
viewer.IsOwnForm = true	  
viewer.IsOwnPage = true

Dim theReportName
theReportName = Request.Form("ReportName")
if theReportName = "" then theReportName = Request.QueryString("ReportName")
viewer.URI = "interactiveViewer.asp?ReportName=" + Server.URLEncode(theReportName)

Dim clientDoc
Set clientDoc = objectFactory.CreateObject("CrystalClientDoc.ReportClientDocument")
clientDoc.Open theReportName
viewer.ReportSource = clientDoc.ReportSource

Dim BooleanSearchControl
Set BooleanSearchControl = objectFactory.CreateObject("CrystalReports.BooleanSearchControl")
BooleanSearchControl.ReportDocument = clientDoc
viewer.BooleanSearchControl = BooleanSearchControl
viewer.ProcessHttpRequest Request, Response, Session
' ReportClientDocument will be automatically closed when clientDoc is released
%>