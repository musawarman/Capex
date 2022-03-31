<%@LANGUAGE="VBSCRIPT"%> 
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="8"
MM_authFailedURL="failed.asp"
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
<!--#include file="../Connections/CapexConn.asp" -->
<html>
<head>
<title>:: Search System Manager ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="../Image/bg.gif">
<table width="750" border="0" align="center">
  <tr> 
    <td width="916" height="20"><img src="../Image/banner2.gif" width="750" height="100"></td>
  </tr>
  <tr> 
    <td height="20"><div align="center"><font color="#0000FF">Date :</font> <%=date %> 
      </div></td>
  </tr>
  <tr> 
    <td height="14"><div align="center"><font color="#FF0000">Welcome</font> <%= Session("UpdateUsr") %></div></td>
  </tr>
  <tr> 
    <td height="14"><div align="center"><strong><font color="#0000FF">To start 
        your search, just click the navigation bar in the left</font></strong></div></td>
  </tr>
  <tr>
    <td height="14"><div align="center">
      <table width="600" border="0">
        <tr>
          <td><div align="center"><a href="../MainMenu/MasterBudget.asp" target="_parent">Master Budget</a> | <a href="../MainMenu/MasterCompany.asp" target="_parent">Master Company</a> | <a href="../MainMenu/MasterCurrency.asp" target="_parent">Master Currency </a> | <a href="../MainMenu/MasterDivisi.asp" target="_parent">Master Divisi </a>| <a href="../MainMenu/MasterUser.asp" target="_parent">Master User </a> | <a href="../MainMenu/MasterVendor.asp" target="_parent">Master Vendor </a></div></td>
        </tr>
      </table>
    </div></td>
  </tr>
</table>
<div align="left"> </div>
</body>
</html>

