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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v6.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}
//-->
</script>
<script language="JavaScript" src="MainMenu/mm_menu.js"></script>
<style type="text/css">
<!--
.style8 {	font-size: x-small;
	color: #000066;
	font-family: Arial, Helvetica, sans-serif;
}
.style9 {
	color: #FFFF00;
	font-weight: bold;
}
a:link {
	color: #FFFFFF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #FFFFFF;
}
a:hover {
	text-decoration: none;
	color: #FF0000;
}
a:active {
	text-decoration: none;
}
body {
	background-image: url(../Image/bg.gif);
}
-->
</style>
</head>

<body onLoad="MM_preloadImages('Image/SysMan.gif','../Image/budget_over.gif','../Image/company_over.gif','../Image/currency_over.gif','../Image/divisi_over.gif','../Image/user_over.gif','../Image/vendor_over.gif')">
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF"> 
    <td height="102" colspan="3"> <img src="../Image/banner2.gif" width="750" height="100"></td>
  </tr>
  <tr> 
    <td width="200" bgcolor="#006666"> <div align="left"><strong><font color="#FFFFFF">Welcome 
        <%= Session("UpdateUsr") %> </font></strong></div></td>
    <td width="423" bgcolor="#006666"><div align="center"><strong><font color="#FFFFFF">..:: <a href="../Search/Search.asp">searching</a> ::.. </font></strong></div></td>
    <td width="127" bgcolor="#006666"><div align="right"><a href="<%= MM_Logout %>"><font color="#FFFFFF" size="-1">Logout</font></a></div></td>
  </tr>
  <tr bgcolor="#FF9900"> 
    <td bgcolor="#660000"><a href="MasterBudget.asp" target="_top" onClick="MM_nbGroup('down','group1','budget','',1)" onMouseOver="MM_nbGroup('over','budget','../Image/budget_over.gif','',1)" onMouseOut="MM_nbGroup('out')"><img src="../Image/budget.gif" alt="" name="budget" width="200" height="50" border="0" onload=""></a><br>
      <a href="MasterCompany.asp" target="_top" onClick="MM_nbGroup('down','group1','company','',1)" onMouseOver="MM_nbGroup('over','company','../Image/company_over.gif','',1)" onMouseOut="MM_nbGroup('out')"><img src="../Image/company.gif" alt="" name="company" width="200" height="50" border="0" onload=""></a><br>
      <a href="MasterCurrency.asp" target="_top" onClick="MM_nbGroup('down','group1','currency','',1)" onMouseOver="MM_nbGroup('over','currency','../Image/currency_over.gif','',1)" onMouseOut="MM_nbGroup('out')"><img src="../Image/currency.gif" alt="" name="currency" width="200" height="50" border="0" onload=""></a><br>
      <a href="MasterDivisi.asp" target="_top" onClick="MM_nbGroup('down','group1','divisi','',1)" onMouseOver="MM_nbGroup('over','divisi','../Image/divisi_over.gif','',1)" onMouseOut="MM_nbGroup('out')"><img src="../Image/divisi.gif" alt="" name="divisi" width="200" height="50" border="0" onload=""></a><br>
      <a href="MasterUser.asp" target="_top" onClick="MM_nbGroup('down','group1','user','',1)" onMouseOver="MM_nbGroup('over','user','../Image/user_over.gif','',1)" onMouseOut="MM_nbGroup('out')"><img src="../Image/user.gif" alt="" name="user" width="200" height="50" border="0" onload=""></a><br>
    <a href="MasterVendor.asp" target="_top" onClick="MM_nbGroup('down','group1','vendor','',1)" onMouseOver="MM_nbGroup('over','vendor','../Image/vendor_over.gif','',1)" onMouseOut="MM_nbGroup('out')"><img src="../Image/vendor.gif" alt="" name="vendor" width="200" height="50" border="0" onload=""></a><br> </td>
    <td height="72" bgcolor="#FF9900"><p align="center"><strong><font color="#0000FF" size="+1">SYSTEM MANAGER</font></strong></p>
    <p align="center"><font size="-1">Click the navigation bars on the left and right to use this modul </font></p></td>
    <td bgcolor="#660000"><p><strong><font color="#FFFFFF">:: <a href="ListBudget.asp">list budget</a> ::</font></strong></p>
    <p><strong><font color="#FFFFFF">:: <a href="ListCompany.asp">list company</a> ::</font></strong></p>
    <p><strong><font color="#FFFFFF">:: <a href="ListCurrency.asp">list currency </a> ::</font></strong></p>
    <p><strong><font color="#FFFFFF">:: <a href="ListDivisi.asp">list divisi</a> :: </font></strong></p>
    <p><strong><font color="#FFFFFF">:: <a href="ListUser.asp">list user</a> ::</font></strong></p>
    <p><strong><font color="#FFFFFF">:: <a href="ListVendor.asp">list vendor</a> ::</font></strong></p></td>
  </tr>
  <tr bgcolor="#006666"> 
    <td colspan="3"><div align="center" class="style9"><font size="-1">Date 
        : 
        <% =date %>
    </font>     </div></td>
  </tr>
  <tr bgcolor="#00CC00"> 
    <td height="22" colspan="3"> <div align="center"> <span class="style8">Copyright &copy; 2005 PT. Sierad Produce Tbk. </span></div></td>
  </tr>
</table>
</body>
</html>
