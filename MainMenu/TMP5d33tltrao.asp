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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Main Page ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style2 {color: #FF0000}
.style3 {
	font-size: 18px;
	color: #0000FF;
}
-->
</style>
<script language="JavaScript" type="text/JavaScript"
>
<!--



function mmLoadMenus() {
  if (window.mm_menu_0711120407_0) return;
    window.mm_menu_0711120407_0 = new Menu("root",96,20,"Geneva, Arial, Helvetica, sans-serif",14,"#FFFFFF","#0000FF","#009900","#FF9900","left","middle",3,0,300,-5,7,true,true,true,0,true,true);
  mm_menu_0711120407_0.addMenuItem("masdewew","location='Master company.asp'");
   mm_menu_0711120407_0.fontWeight="bold";
   mm_menu_0711120407_0.hideOnMouseOut=true;
   mm_menu_0711120407_0.bgColor='#00FF66';
   mm_menu_0711120407_0.menuBorder=1;
   mm_menu_0711120407_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0711120407_0.menuBorderBgColor='#CC3300';

  window.mm_menu_0711113145_0 = new Menu("root",134,20,"Geneva, Arial, Helvetica, sans-serif",14,"#FFFFFF","#0000FF","#009900","#FF9900","left","middle",3,0,300,-5,7,true,true,true,0,true,true);
  mm_menu_0711113145_0.addMenuItem("Master&nbsp;Company","location='Master company.asp'");
   mm_menu_0711113145_0.fontWeight="bold";
   mm_menu_0711113145_0.hideOnMouseOut=true;
   mm_menu_0711113145_0.bgColor='#00FF66';
   mm_menu_0711113145_0.menuBorder=1;
   mm_menu_0711113145_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0711113145_0.menuBorderBgColor='#CC3300';

  window.mm_menu_0711113145_0 = new Menu("root",134,20,"Geneva, Arial, Helvetica, sans-serif",14,"#FFFFFF","#0000FF","#009900","#FF9900","left","middle",3,0,300,-5,7,true,true,true,0,true,true);
  mm_menu_0711113145_0.addMenuItem("Master&nbsp;Company","window.open('Master company.asp', '_blank');");
   mm_menu_0711113145_0.fontWeight="bold";
   mm_menu_0711113145_0.hideOnMouseOut=true;
   mm_menu_0711113145_0.bgColor='#00FF66';
   mm_menu_0711113145_0.menuBorder=1;
   mm_menu_0711113145_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0711113145_0.menuBorderBgColor='#CC3300';

  window.mm_menu_0711113145_0 = new Menu("root",138,20,"Geneva, Arial, Helvetica, sans-serif",14,"#FFFFFF","#0000FF","#009900","#FF9900","left","middle",3,0,300,-5,7,true,true,true,0,true,true);
  mm_menu_0711113145_0.addMenuItem("MAster&nbsp;Company","location='Master company.asp'");
   mm_menu_0711113145_0.fontWeight="bold";
   mm_menu_0711113145_0.hideOnMouseOut=true;
   mm_menu_0711113145_0.bgColor='#00FF66';
   mm_menu_0711113145_0.menuBorder=1;
   mm_menu_0711113145_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0711113145_0.menuBorderBgColor='#CC3300';

    window.mm_menu_0712092225_0 = new Menu("root",89,20,"Geneva, Arial, Helvetica, sans-serif",14,"#FFFFFF","#0000FF","#009900","#FF9900","left","middle",3,0,300,-5,7,true,true,true,0,true,true);
  mm_menu_0712092225_0.addMenuItem("New&nbsp;Item");
   mm_menu_0712092225_0.fontWeight="bold";
   mm_menu_0712092225_0.hideOnMouseOut=true;
   mm_menu_0712092225_0.bgColor='#00FF66';
   mm_menu_0712092225_0.menuBorder=1;
   mm_menu_0712092225_0.menuLiteBgColor='#FFFFFF';
   mm_menu_0712092225_0.menuBorderBgColor='#CC3300';

mm_menu_0712092225_0.writeMenus();
} // mmLoadMenus()<!--




<!--

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->

//-->

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
</script>
<script language="JavaScript" src="mm_menu.js"></script>
</head>

<body background="../Image/bgco.gif" onLoad="MM_preloadImages('../Image/SysMan.gif')">
<table width="600" border="0" align="center" cellspacing="0">
  <tr> 
    <td height="97" colspan="3">&nbsp;</td>
  </tr>
  <tr bgcolor="#006666"> 
    <td height="102" colspan="3"> <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="600" height="100" align="absmiddle">
        <param name="movie" value="../Animasi/baner.swf">
        <param name="SCALE" value="exactfit">
        <param name="BGCOLOR" value="#006666">
        <embed src="../Animasi/baner.swf" width="600" height="100" align="absmiddle" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit" bgcolor="#006666"></embed></object></td>
  </tr>
  <tr> 
    <td bgcolor="#006666"> <div align="left"><strong><font color="#FFFFFF">Welcome 
        <%= Session("UpdateUsr") %> </font></strong></div></td>
    <td bgcolor="#006666">&nbsp;</td>
    <td bgcolor="#006666"> <div align="right"><a href="<%= MM_Logout %>"><font color="#FFFFFF" size="-1">Logout</font></a></div></td>
  </tr>
  <tr bgcolor="#FF9900"> 
    <td height="72" colspan="3"> <div align="center"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="150" height="70">
          <param name="movie" value="ActButtn.swf">
          <param name="quality" value="high">
          <param name="base" value=".">
          <param name="BGCOLOR" value="#FF9900">
          <embed src="ActButtn.swf" width="150" height="70" base="."  quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#FF9900" ></embed> 
        </object>
      </div></td>
  </tr>
  <tr> 
    <td width="400" bgcolor="#FF9900"> <div align="center"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="150" height="70">
          <param name="BASE" value=".">
          <param name="movie" value="SysmanBttn.swf">
          <param name="quality" value="high">
          <param name="BGCOLOR" value="#FF9900">
          <embed src="SysmanBttn.swf" width="150" height="70" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" bgcolor="#FF9900" base="."></embed></object>
      </div></td>
    <td width="400" bgcolor="#FF9900"><div align="center"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="120" height="50">
          <param name="BASE" value=".">
          <param name="BGCOLOR" value="#FF9900">
          <param name="movie" value="AboutBttn.swf">
          <param name="quality" value="high">
          <embed src="AboutBttn.swf" width="120" height="50" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#FF9900" base="." ></embed> 
        </object>
      </div></td>
    <td width="400" bgcolor="#FF9900"> <div align="center"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="150" height="70">
          <param name="movie" value="ReportBttn.swf">
          <param name="quality" value="high">
          <param name="bgcolor" value="#FF9900">
          <embed src="ReportBttn.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="150" height="70" bgcolor="#FF9900"></embed> 
        </object>
      </div></td>
  </tr>
  <tr bgcolor="#FF9900"> 
    <td colspan="3"> <div align="center"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="150" height="70">
          <param name="movie" value="HelpBttn.swf">
          <param name="quality" value="high">
          <param name="bgcolor" value="#FF9900">
          <embed src="HelpBttn.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="150" height="70" bgcolor="#FF9900"></embed> 
        </object>
      </div></td>
  </tr>
  <tr bgcolor="#00FF00"> 
    <td colspan="3"><div align="center"><font color="#990000" size="-1"><strong>Date 
        : 
        <% =date %>
        </strong> </font></div></td>
  </tr>
  <tr bgcolor="#00FF00"> 
    <td height="22" colspan="3"> <div align="center"> <strong><font color="#0000FF" size="4">PT</font><font color="#006600" size="4"> 
        Sierad<font color="#FF9900">Produce</font> <font color="#FF0000">Tbk.</font></font></strong></div></td>
  </tr>
</table>
</body>
</html>
