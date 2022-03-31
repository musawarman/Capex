<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "login.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<!--#include file="../Connections/CapexConn.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="8"
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>

<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_CapexConn_STRING
  MM_editTable = "dbo.CategoryBudget"
  MM_editColumn = "BudgetID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "ListBudget.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsBudget__MMColParam
rsBudget__MMColParam = "1"
If (Request.QueryString("BudgetID") <> "") Then 
  rsBudget__MMColParam = Request.QueryString("BudgetID")
End If
%>
<%
Dim rsBudget
Dim rsBudget_numRows

Set rsBudget = Server.CreateObject("ADODB.Recordset")
rsBudget.ActiveConnection = MM_CapexConn_STRING
rsBudget.Source = "SELECT * FROM dbo.CategoryBudget WHERE BudgetID = '" + Replace(rsBudget__MMColParam, "'", "''") + "'"
rsBudget.CursorType = 0
rsBudget.CursorLocation = 2
rsBudget.LockType = 1
rsBudget.Open()

rsBudget_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Delete Budget Confirmation ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-size: 18px}
.style4 {font-size: 14px}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style5 {color: #6600FF}
.style10 {
	font-family: verdana;
	font-weight: bold;
}
-->
</style>
<% session.timeout=5%>
</head>

<body background="../Image/bg.gif">
<div align="center"> 
  <table width="600" border="0">
    <tr> 
      <td colspan="2"><img src="../Image/banner2.gif" width="750" height="100"></td>
    </tr>
    <tr> 
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: Confirmation 
          ::. </font></h3></td>
    </tr>
    <tr> 
      <td width="468"><div align="left"> 
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="50" height="37" hspace="10" align="absmiddle" onMouseOver="MM_displayStatusMsg('rz : delete vendor confirmation -&gt; back to system manager');return document.MM_returnValue">
            <param name="BASE" value=".">
            <param name="BGCOLOR" value="">
            <param name="movie" value="back2sm.swf">
            <embed src="back2sm.swf" width="50" height="37" hspace="10" align="absmiddle" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" base="." ></embed> 
          </object>
          <span class="style4 style5"><font color="#0000FF" size="-1" face="Arial, Helvetica, sans-serif">back 
          to system manager </font></span> </div></td>
      <td width="120">
<div align="center"><font size="2"><a href="<%= MM_Logout %>" target="_parent">Logout</a> 
          |<strong> <a href="../Search/Search.asp" target="_parent">Search</a></strong></font></div></td>
    </tr>
  </table>
  <div align="left"> </div>
  <table width="600" border="1">
    <tr>
      <td width="248" height="23" bgcolor="#669900">
<div align="left"><font color="#FFFFFF">-- 
          CONFIRM DELETE BUDGET --</font></div></td>
    </tr>
  </table>
  <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
    <table width="553" border="1" cellspacing="0" bordercolor="#999999">
      <tr bgcolor="#000099"> 
        <td colspan="2"> 
          <div align="center"> <font color="#FFFFFF" size="3" face="Comic Sans MS">Budget 
            Category Information</font></div>          </td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right" class="bhs2  style10"><font color="#FF0000"><strong>Budget 
            ID :</strong></font> </div></td>
        <td width="330"> 
          <div align="left"><strong><%=(rsBudget.Fields.Item("BudgetID").Value)%></strong></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right"><strong><font color="#FF0000">Budget Name : </font> </strong></div></td>
        <td> 
          <div align="left"><strong><%=(rsBudget.Fields.Item("BudgetName").Value)%></strong></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right"><strong><font color="#FF0000">Total Budget</font> : </strong></div></td>
        <td> 
          <div align="left"><strong><%=(rsBudget.Fields.Item("TotalBudget").Value)%></strong></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right"><strong><font color="#FF0000">Tanggal Budget : </font></strong></div></td>
        <td> 
          <div align="left"><strong><%=(rsBudget.Fields.Item("TanggalBudget").Value)%></strong></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right"><strong><font color="#FF0000">Saldo Budget : </font></strong></div></td>
        <td> 
          <div align="left"><strong><%=(rsBudget.Fields.Item("SaldoBudget").Value)%></strong></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td height="28"> 
          <input name="Submit" type="submit" onMouseOver="MM_displayStatusMsg('rz : delete budget confirmation -&gt; delete record');return document.MM_returnValue" value="Delete"> </td>
        <td> 
          <div align="right"><font color="#0000FF">Date:</font> <%=date%></div></td>
      </tr>
    </table>
  
    <input type="hidden" name="MM_delete" value="form1">
    <input type="hidden" name="MM_recordId" value="<%= rsBudget.Fields.Item("BudgetID").Value %>">
</form>
  <form name="form2" method="post" action="">
    <input name="Submit2" type=button class=btn onclick=history.back() onMouseOver="MM_displayStatusMsg('rz : delete company confirmation -&gt; back to last form');return document.MM_returnValue" value="Cancel">
  </form>
  <table width="600" border="1">
    <tr>
      <td bordercolor="#FFFFCC" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <br>
  <table width="600" border="0">
    <tr> 
      <td><div align="center"><a href="MasterBudget.asp" target="_parent">Master 
          Budget</a> | <a href="MasterCompany.asp" target="_parent">Master 
          Company</a> | <a href="MasterCurrency.asp" target="_parent">Master Currency </a> 
          | <a href="MasterDivisi.asp" target="_parent">Master Divisi </a>| <a href="MasterUser.asp" target="_parent">Master User </a> | 
          <a href="MasterVendor.asp" target="_parent">Master Vendor </a></div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsBudget.Close()
Set rsBudget = Nothing
%>
