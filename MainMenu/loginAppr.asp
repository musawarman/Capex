<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../Connections/CapexConn.asp" -->
<%
Dim rsUser
Dim rsUser_numRows

Set rsUser = Server.CreateObject("ADODB.Recordset")
rsUser.ActiveConnection = MM_CapexConn_STRING
rsUser.Source = "SELECT * FROM dbo.UserMS"
rsUser.CursorType = 0
rsUser.CursorLocation = 2
rsUser.LockType = 1
rsUser.Open()

rsUser_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("textfield"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization=""
  MM_redirectLoginSuccess="Infocapexappr.asp"
  MM_redirectLoginFailed="loginAppr.asp"
  MM_flag="ADODB.Recordset"
  set MM_rsUser = Server.CreateObject(MM_flag)
  MM_rsUser.ActiveConnection = MM_CapexConn_STRING
  MM_rsUser.Source = "SELECT JabatanID, UserPassword"
  If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
  MM_rsUser.Source = MM_rsUser.Source & " FROM dbo.UserMS WHERE JabatanID='" & Replace(MM_valUsername,"'","''") &"' AND UserPassword='" & Replace(Request.Form("textfield2"),"'","''") & "'"
  MM_rsUser.CursorType = 0
  MM_rsUser.CursorLocation = 2
  MM_rsUser.LockType = 3
  MM_rsUser.Open
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
	session("UpdateJabatan") = Session("MM_Username")
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

<body>
<div id="Layer1" style="position:absolute; left:186px; top:313px; width:757px; height:128px; z-index:1"> 
  <form ACTION="<%=MM_LoginAction%>" name="form1" method="POST" >
    <table width="736" height="177" border="1" align="center">
      <tr bgcolor="#FF3300"> 
        <td width="202" height="171" bordercolor="#3300FF" bgcolor="#FF9900"> 
          <p align="left" class="style2"> <font color="#000000">Jabatan ID</font> 
            * 
            <input name="textfield" type="text" value="">
          </p>
          <p align="left" class="style2"> <font color="#000000">Password</font> 
            * 
            <input type="password" name="textfield2">
          </p>
          <p align="left" class="style2"> 
            <input type="submit" name="Login" value="Login">
            <input name="hiddenField" type="hidden" value="<%= Session("UpdateUsr") %>">
            <input name="hidefield" type="hidden" id="hidefield" value="<%= Session("UpdateJabatan") %>">
          </p></td>
        <td width="636" colspan="4" bgcolor="#FF9900"><p align="center" class="style3">CapeX 
            OnLine</p>
          <div align="center">Capital Expenditure Intergrated System Online(CapexOL) 
            Sistem informasi yang hanya diperuntukkan bagi pihak-pihak tertentu 
            yang terkait dengan <span class="style5">PT.SieradProduce Tbk.</span></div></td>
      </tr>
    </table>
  </form>
</div>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="1100" height="153" align="absbottom" class="trdark">
  <param name="movie" value="../Animasi/baner.swf">
  <param name="quality" value="high">
  <param name="SCALE" value="exactfit">
  <embed src="../Animasi/baner.swf" width="1100" height="153" align="absbottom" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
<div align="center"> 
  <table width="1100" border="1" align="center">
    <tr> 
      <td width="997" bgcolor="#669900" class="style2">&nbsp;</td>
    </tr>
    <tr> 
      <td height="72" bgcolor="#FF9900"> <p align="center" class="style6"> 
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="1090" height="34" align="left">
            <param name="movie" value="../Animasi/Texttitle.swf">
            <param name="quality" value="high">
            <param name="SCALE" value="exactfit">
            <embed src="../Animasi/Texttitle.swf" width="1090" height="34" align="left" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
        </p></td>
    </tr>
    <tr> 
      <td bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp; </p>
  <p>&nbsp;</p>
  <p align="justify">&nbsp;</p>
  <p align="justify">&nbsp;</p>
  <p align="justify">&nbsp; </p>
  <table width="1077" border="1">
    <tr> 
      <td width="1032" height="202" bgcolor="#339900"> <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p></td>
    </tr>
  </table>
  
</div>
</body>
</html>
<%
rsUser.Close()
Set rsUser = Nothing
%>
