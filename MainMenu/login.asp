<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="../Connections/CapexConn.asp" -->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("textfield"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="UserLEvel"
  MM_redirectLoginSuccess="mainpage.asp"
  MM_redirectLoginFailed="login.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_CapexConn_STRING
  MM_rsUser.Source = "SELECT UserID, UserPassword"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.UserMS WHERE UserID='" & Replace(MM_valUsername,"'","''") &"' AND UserPassword='" & Replace(Request.Form("textfield2"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
	Session("UpdateUsr") = session("MM_Username")
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Login ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style2 {font-size: 18px}
.style3 {font-size: 24px}
body,td,th {
	color: #0000FF;
}
.style5 {font-size: 24px; color: #33FF00; }
.style6 {font-size: 36px}
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
//-->
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#CCCCCC">
<table width="800" height="110" border="1" align="center" bgcolor="#999900">
  <tr>
    <td height="104"> 
      <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="800" height="100" align="top" class="trdark">
        <param name="movie" value="../Animasi/baner.swf">
        <param name="quality" value="high">
        <param name="SCALE" value="exactfit">
        <embed src="../Animasi/baner.swf" width="800" height="100" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object></td>
  </tr>
</table>
<div align="center"> 
  <p>&nbsp;</p>
  <table width="719" border="0">
    <tr> 
      <td width="3">&nbsp;</td>
      <td width="700" height="3"><div align="left"><font color="#660033">HALAMAN LOGIN </font></div></td>
      <td width="10">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="3"> <div align="center"><img src="../../img/garis1.gif" width="700" height="1"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td><img src="../../img/garis.gif" width="1" height="300"></td>
      <td height="150"> <form name="form1" method="POST" action="<%=MM_LoginAction%>" >
          <table width="700" height="148" border="1" align="center" bordercolor="#FF9900" bgcolor="#6699FF">
            <tr> 
              <td width="202" height="142"> <p align="left" class="style2"> <font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">User 
                  Name</font> 
                  <input name="textfield" type="text" value="">
                </p>
                <p align="left" class="style2"> <font color="#FFFF00" size="2">Password</font> 
                  <input type="password" name="textfield2">
                </p>
                <p align="center" class="style2"> 
                  <input type="submit" name="Login" value="Login">
                  <input name="hiddenField" type="hidden" value="<%= Session("UpdateUsr") %>">
                  <input name="hidefield" type="hidden" id="hidefield" value="<%= Session("UpdateJabatan") %>">
                  <input type="reset" name="Reset" value="Reset">
                </p></td>
              <td width="636" colspan="4"> <p align="center" class="style3"><font color="#FFFFFF" size="4">CapeX 
                  OnLine</font></p>
                <div align="center"><font color="#FFFFFF">Capital Expenditure 
                  Intergrated System Online(CapexOL) Sistem informasi yang hanya 
                  diperuntukkan bagi pihak-pihak tertentu yang terkait dengan</font> 
                  <span class="style5">PT.SieradProduce Tbk.</span></div></td>
            </tr>
          </table>
        </form></td>
      <td><img src="../../img/garis.gif" width="1" height="300"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="3"><img src="../../img/garis1.gif" width="700" height="1"></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  
</div>
</body>
</html>
