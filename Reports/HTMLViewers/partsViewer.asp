<%@CodePage=65001%>
<% 
Option Explicit
Response.ContentType = "text/html; charset=utf-8"

Dim objectFactory
Set objectFactory = CreateObject("CrystalReports.ObjectFactory.2")

Response.ExpiresAbsolute = Now() - 1
	
Dim viewer
Set viewer = objectFactory.CreateObject("CrystalReports.CrystalReportPartsViewer")  
viewer.Name = "parts"
viewer.IsOwnForm = true	  
viewer.IsOwnPage = true

Dim theReportName
theReportName = Request.QueryString("ReportName")
if theReportName <> "" then viewer.ReportSource = theReportName
viewer.ProcessHttpRequest Request, Response, Session

%>