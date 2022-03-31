<%@LANGUAGE="VBSCRIPT"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("MM_Username")
  Session.Contents.Remove("MM_UserAuthorization")
  MM_logoutRedirectPage = "../MainMenu/login.asp"
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
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="8"
MM_authFailedURL="failed2.asp"
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>:: Search System Manager ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function mmLoadMenus() {
  if (window.mm_menu_0806111527_0) return;
          window.mm_menu_0806111527_0 = new Menu("root",106,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806111527_0.addMenuItem("Budget&nbsp;ID","window.open('SearchBudgetID.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Budget&nbsp;Name","window.open('SearchBudgetName.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Total&nbsp;Budget","window.open('SearchBudTotal.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Saldo&nbsp;Budget","window.open('SearchBudSaldo.asp', 'mainFrame');");
  mm_menu_0806111527_0.addMenuItem("Post&nbsp;By","window.open('SearchBudPostBy.asp', 'mainFrame');");
   mm_menu_0806111527_0.fontWeight="bold";
   mm_menu_0806111527_0.hideOnMouseOut=true;
   mm_menu_0806111527_0.bgColor='#555555';
   mm_menu_0806111527_0.menuBorder=1;
   mm_menu_0806111527_0.menuLiteBgColor='';
   mm_menu_0806111527_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112043_0 = new Menu("root",122,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112043_0.addMenuItem("Company&nbsp;ID","window.open('SearchCompID.asp', 'mainFrame');");
  mm_menu_0806112043_0.addMenuItem("Company&nbsp;Name","window.open('SearchCompName.asp', 'mainFrame');");
  mm_menu_0806112043_0.addMenuItem("Post&nbsp;By","window.open('SearchCompPostBy.asp', 'mainFrame');");
   mm_menu_0806112043_0.fontWeight="bold";
   mm_menu_0806112043_0.hideOnMouseOut=true;
   mm_menu_0806112043_0.bgColor='#555555';
   mm_menu_0806112043_0.menuBorder=1;
   mm_menu_0806112043_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112043_0.menuBorderBgColor='#FFFF00';
window.mm_menu_0806112209_0 = new Menu("root",119,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112209_0.addMenuItem("Currency&nbsp;ID","window.open('SearchCurrID.asp', 'mainFrame');");
  mm_menu_0806112209_0.addMenuItem("Currency&nbsp;Name","window.open('SearchCurrName.asp', 'mainFrame');");
  mm_menu_0806112209_0.addMenuItem("Post&nbsp;By","window.open('SearchCurrPostBy.asp', 'mainFrame');");
   mm_menu_0806112209_0.fontWeight="bold";
   mm_menu_0806112209_0.hideOnMouseOut=true;
   mm_menu_0806112209_0.bgColor='#555555';
   mm_menu_0806112209_0.menuBorder=1;
   mm_menu_0806112209_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112209_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112240_0 = new Menu("root",101,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112240_0.addMenuItem("Divisi&nbsp;ID","window.open('SearchDivID.asp', 'mainFrame');");
  mm_menu_0806112240_0.addMenuItem("Company&nbsp;ID","window.open('SearchDivComID.asp', 'mainFrame');");
  mm_menu_0806112240_0.addMenuItem("Divisi&nbsp;Name","window.open('SearchDivName.asp', 'mainFrame');");
  mm_menu_0806112240_0.addMenuItem("Post&nbsp;By","window.open('SearchDivPostBy.asp', 'mainFrame');");
   mm_menu_0806112240_0.fontWeight="bold";
   mm_menu_0806112240_0.hideOnMouseOut=true;
   mm_menu_0806112240_0.bgColor='#555555';
   mm_menu_0806112240_0.menuBorder=1;
   mm_menu_0806112240_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112240_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112351_0 = new Menu("root",92,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112351_0.addMenuItem("Jabatan","window.open('SearchUsrJabatan.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;ID","window.open('SearchUsrID.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;Name","window.open('SearchUsrName.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;Level","window.open('SearchUsrLevel.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("User&nbsp;Status","window.open('SearchUsrStat.asp', 'mainFrame');");
  mm_menu_0806112351_0.addMenuItem("Post&nbsp;By","window.open('SearchUsrPostBy.asp', 'mainFrame');");
   mm_menu_0806112351_0.fontWeight="bold";
   mm_menu_0806112351_0.hideOnMouseOut=true;
   mm_menu_0806112351_0.bgColor='#555555';
   mm_menu_0806112351_0.menuBorder=1;
   mm_menu_0806112351_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112351_0.menuBorderBgColor='#FFFF00';
  window.mm_menu_0806112648_0 = new Menu("root",116,17,"",11,"#FFFF00","#FF0000","#330066","#33FF99","left","middle",3,0,500,-5,7,true,true,true,0,true,true);
  mm_menu_0806112648_0.addMenuItem("Vendor&nbsp;ID","window.open('SearchVendorID.asp', 'mainFrame');");
  mm_menu_0806112648_0.addMenuItem("Vendor&nbsp;Name","window.open('SearchVenName.asp', 'mainFrame');");
  mm_menu_0806112648_0.addMenuItem("Contact&nbsp;Person","window.open('SearchVenCp.asp', 'mainFrame');");
  mm_menu_0806112648_0.addMenuItem("Post&nbsp;By","window.open('SearchVenPostBy.asp', 'mainFrame');");
   mm_menu_0806112648_0.fontWeight="bold";
   mm_menu_0806112648_0.hideOnMouseOut=true;
   mm_menu_0806112648_0.bgColor='#555555';
   mm_menu_0806112648_0.menuBorder=1;
   mm_menu_0806112648_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0806112648_0.menuBorderBgColor='#FFFF00';

mm_menu_0806112648_0.writeMenus();
} // mmLoadMenus()

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
<script language="JavaScript" src="mm_menu.js"></script>
</head>

<body bgcolor="#330066" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<p> 
  <script language="JavaScript1.2">mmLoadMenus();</script>
</p>
<a href="<%= MM_Logout %>" target="_parent"><font size="-1" face="Courier New, Courier, mono">Logout</font></a> 
<p> 
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="107" height="18" onMouseOver="MM_displayStatusMsg('rz : search -&gt; back to home');return document.MM_returnValue">
    <param name="movie" value="home.swf">
    <param name="quality" value="high">
    <param name="base" value=".">
    <param name="BGCOLOR" value="#330066">
    <embed src="home.swf" width="107" height="18" base="."  quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#330066" ></embed> 
  </object>
</p>
<p align="center"><font color="#FFFF00"><strong>Search By<br>
  _________________ </strong></font></p>
<p><img src="../Image/ListBudget.gif" name="listbudget" width="125" height="25" id="listbudget" onMouseOver="MM_showMenu(window.mm_menu_0806111527_0,0,25,null,'listbudget');MM_displayStatusMsg('rz : search by list budget');return document.MM_returnValue" onMouseOut="MM_startTimeout();"> 
  <br>
  <br>
  <img src="../Image/ListCompany.gif" name="listcompany" width="125" height="25" id="listcompany" onMouseOver="MM_showMenu(window.mm_menu_0806112043_0,0,25,null,'listcompany');MM_displayStatusMsg('rz : search by list company');return document.MM_returnValue" onMouseOut="MM_startTimeout();"><br>
  <br>
  <img src="../Image/ListCurrency.gif" name="listcurrency" width="125" height="25" id="listcurrency" onMouseOver="MM_showMenu(window.mm_menu_0806112209_0,0,25,null,'listcurrency');MM_displayStatusMsg('rz : search by list currency');return document.MM_returnValue" onMouseOut="MM_startTimeout();"><br>
  <br>
  <img src="../Image/ListDivisi.gif" name="listdivisi" width="125" height="25" id="listdivisi" onMouseOver="MM_showMenu(window.mm_menu_0806112240_0,0,25,null,'listdivisi');MM_displayStatusMsg('rz : search by list divisi');return document.MM_returnValue" onMouseOut="MM_startTimeout();"> 
  <br>
  <br>
  <img src="../Image/ListUser.gif" name="listuser" width="125" height="25" id="listuser" onMouseOver="MM_showMenu(window.mm_menu_0806112351_0,0,25,null,'listuser');MM_displayStatusMsg('rz : search by list user');return document.MM_returnValue" onMouseOut="MM_startTimeout();"> 
  <br>
  <br>
  <img src="../Image/ListVendor.gif" name="listvendor" width="125" height="25" id="listvendor" onMouseOver="MM_showMenu(window.mm_menu_0806112648_0,0,25,null,'listvendor');MM_displayStatusMsg('rz : search by list vendor');return document.MM_returnValue" onMouseOut="MM_startTimeout();"><br>
  <font color="#FFFF00">_________________</font></p>
<p>&nbsp; </p>
<p>&nbsp;</p>
</body>
</html>
