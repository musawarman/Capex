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
<html>
<head>
<title>Activity</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/stylee.css" rel="stylesheet" type="text/css">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body background="../Image/bgact.gif" topmargin="0" onLoad="MM_preloadImages('../Image/home_capex_on.jpg','../Image/sysman_capex_on.jpg','../Image/activities_capex_on.jpg','../Image/report_capex_on.jpg','../Image/faq_capex_on.jpg','../Image/CreateCapex_on.gif','../Image/ApprovalCapex_on.gif','../Image/CreateAOC_on.gif','../Image/Approvalaoc_on.gif','../Image/holdcapex_on.gif','../Image/holdAoc_on.gif','../Image/Estimation_on.gif','../Image/Actual_on.gif')">
<table width="800" border="1" align="center" bordercolor="#006600">
  <tr bordercolor="#006600"> 
    <td colspan="2"> <div align="left"><img src="../Image/sieradonline.gif" width="222" height="85"> 
      </div>
      <div align="right"><font color="#006600">Date : 
        <script name="current" src="../../GeneratedItems/current.js" language="JavaScript1.2"></script>
        </font></div></td>
  </tr>
  <tr bordercolor="#006600" bgcolor="#CCCCCC"> 
    <td width="487"> <div align="left"><font color="#006600">Welcome <%= Session("UpdateUsr") %></font></div></td>
    <td width="297"> <div align="center"><font color="#009900"><a href="../contact.asp"><font color="#006600">Hubungi 
        Kami</font></a></font><font color="#FF0000">&nbsp; </font>| <a href="../karir.asp"><font color="#006600">Karir 
        </font></a>| <a href="../link.asp"><font color="#006600">Links </font></a>| 
        <a href="<%= MM_Logout %>"><font color="#006600">Log Out</font></a></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
  <tr> 
    <td><div align="center"><img src="../../BREEDING/img/spacer.gif" width="795" height="10"></div></td>
  </tr>
</table>
<table width="806" border="2" align="center" bordercolor="#CCCCCC" bgcolor="#006600">
  <tr> 
    <td width="150" height="23"><div align="center"><a href="MainPage.asp" onMouseOver="MM_swapImage('Image1','','../Image/home_capex_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/home_capex.jpg" name="Image1" width="150" height="20" border="0" id="Image1"></a></div></td>
    <td width="330" rowspan="5" bgcolor="#006600"><div align="center"> 
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="323" height="122" align="middle">
          <param name="movie" value="../Animasi/anakayam.swf">
          <param name="quality" value="high">
          <param name="SCALE" value="exactfit">
          <embed src="../Animasi/anakayam.swf" width="323" height="122" align="middle" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
      </div></td>
    <td width="296"><div align="right"><font color="#FFFFFF">Time : <strong><%=time()%></strong></font></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="Sman.asp" onMouseOver="MM_swapImage('Image2','','../Image/sysman_capex_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/sysman_capex.jpg" name="Image2" width="150" height="20" border="0" id="Image2"></a></div></td>
    <td rowspan="4"><div align="right"><font color="#FFFFFF"></font></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="Activity.asp" onMouseOver="MM_swapImage('Image3','','../Image/activities_capex_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/activities_capex.jpg" name="Image3" width="150" height="20" border="0" id="Image3"></a></div></td>
  </tr>
  <tr> 
    <td><div align="center"><a href="../Reports/reportListing.asp" onMouseOver="MM_swapImage('Image4','','../Image/report_capex_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/report_capex.jpg" name="Image4" width="150" height="20" border="0" id="Image4"></a></div></td>
  </tr>
  <tr> 
    <td height="24"> <div align="center"><a href="#" onMouseOver="MM_swapImage('Image5','','../Image/faq_capex_on.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/faq_capex.jpg" name="Image5" width="150" height="20" border="0" id="Image5"></a></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
  <tr> 
    <td><div align="center"><img src="../../BREEDING/img/spacer.gif" width="795" height="10"></div></td>
  </tr>
</table>
<table width="800" border="0" align="center" background="../../img/bg.gif">
  <tr> 
    <td width="800" height="20"><font color="#FF0000" size="3" face="Courier New, Courier, mono">&gt;&gt; 
      <strong>ACTIVITIES </strong></font></td>
  </tr>
  <tr> 
    <td><img src="../Image/spacer.jpg" width="200" height="2"></td>
  </tr>
  <tr> 
    <td height="215"> 
      <div align="center"> 
        <table width="800" border="0">
          <tr> 
            <td height="25"><div align="center"><img src="../../img/garis1.gif" width="600" height="1"></div></td>
          </tr>
        </table>
        <p><font color="#009933">-- MODUL --</font></p>
        <p><img src="../../img/garis1.gif" width="400" height="1"></p>
        <table width="400" border="0">
          <tr> 
            <td><div align="center"><a href="createcapex.asp" onMouseOver="MM_swapImage('Image6','','../Image/CreateCapex_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/CreateCapex.gif" name="Image6" width="150" height="23" border="0" id="Image6"></a></div></td>
            <td><div align="center"><a href="Infocapexappr.asp" onMouseOver="MM_swapImage('Image7','','../Image/ApprovalCapex_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/ApprovalCapex.gif" name="Image7" width="150" height="23" border="0" id="Image7"></a></div></td>
          </tr>
          <tr> 
            <td><div align="center"><a href="ListCapexApproved.asp" onMouseOver="MM_swapImage('Image8','','../Image/CreateAOC_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/CreateAOC.gif" name="Image8" width="150" height="23" border="0" id="Image8"></a></div></td>
            <td><div align="center"><a href="InfoAOCappr.asp" onMouseOver="MM_swapImage('Image9','','../Image/Approvalaoc_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Approvalaoc.gif" name="Image9" width="150" height="23" border="0" id="Image9"></a></div></td>
          </tr>
          <tr> 
            <td><div align="center"><a href="hold%20capex.asp" onMouseOver="MM_swapImage('Image10','','../Image/holdcapex_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/holdcapex.gif" name="Image10" width="150" height="23" border="0" id="Image10"></a></div></td>
            <td><div align="center"><a href="hold%20AOC.asp" onMouseOver="MM_swapImage('Image11','','../Image/holdAoc_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/holdAoc.gif" name="Image11" width="150" height="23" border="0" id="Image11"></a></div></td>
          </tr>
          <tr> 
            <td><div align="center"><a href="SearchNoAOCforPayment.asp" onMouseOver="MM_swapImage('Image12','','../Image/Estimation_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Estimation.gif" name="Image12" width="150" height="23" border="0" id="Image12"></a></div></td>
            <td><div align="center"><a href="Payman%20actual.asp" onMouseOver="MM_swapImage('Image13','','../Image/Actual_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Actual.gif" name="Image13" width="150" height="23" border="0" id="Image13"></a></div></td>
          </tr>
        </table>
        <p><img src="../../img/garis1.gif" width="400" height="1"></p>
        <p>&nbsp;</p>
      </div></td>
  </tr>
  <tr>
    <td><div align="center"><img src="../../img/garis1.gif" width="600" height="1"></div></td>
  </tr>
  <tr> 
    <td height="47"><div align="center"> 
        <p>&nbsp;</p><table width="800" border="1" bordercolor="#FF6600">
          <tr> 
            <td><div align="center"><font color="#009900"><font color="#006600">-- 
                HOME</font></font><font color="#006600">&nbsp; </font>| <a href="Sman.asp"><font color="#006600">SYSTEM 
                MANAGER</font></a><font color="#006600"> </font>| <a href="Activity.asp"><font color="#006600">ACTIVITIES</font></a><font color="#006600"> 
                </font>|<a href="../Reports/reportListing.asp"> <font color="#006600">REPORT</font></a><font color="#006600"> 
                | FAQ --</font></div></td>
          </tr>
        </table>
        <p><img src="../Image/bannerrg.jpg" width="800" height="20"></p>
      </div></td>
  </tr>
</table>
</body>
</html>
