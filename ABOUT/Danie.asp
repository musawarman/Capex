<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>:: About Us ::</title>
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

<body background="cat.gif" onLoad="MM_preloadImages('../Image/homee.gif','../Image/aboutus.gif','../Image/contact.gif','../Image/Promotion.gif','../Image/Solutions.gif')">
<table width="740" height="572" border="1" bordercolor="#0000FF">
  <tr> 
    <td height="36" colspan="5"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="753" height="100" align="middle">
        <param name="movie" value="../Animasi/baner.swf">
        <param name="quality" value="high">
        <param name="SCALE" value="exactfit">
        <embed src="../Animasi/baner.swf" width="753" height="100" align="middle" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object></td>
  </tr>
  <tr> 
    <td width="3" height="44" bordercolor="#3366FF">&nbsp;</td>
    <td colspan="2" bordercolor="#3366FF"> <div align="center">Welcome <%= Session("UpdateUsr") %></div></td>
    <td colspan="2" bordercolor="#3366FF"> <div align="left"><a href="../MainMenu/Activity.asp" onMouseOver="MM_swapImage('onhome','','../Image/homee.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/homee_on.gif" name="onhome" width="83" height="42" border="0" id="onhome"></a><a href="about.asp" onMouseOver="MM_swapImage('onabout','','../Image/aboutus.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/aboutus_on.gif" name="onabout" width="95" height="42" border="0" id="onabout"></a><a href="../MainMenu/Contacts.asp" onMouseOver="MM_swapImage('oncontact','','../Image/contact.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/contact_on.gif" name="oncontact" width="95" height="42" border="0" id="oncontact"></a><a href="../MainMenu/Promotion.asp" onMouseOver="MM_swapImage('onprom','','../Image/Promotion.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Promotion_on.gif" name="onprom" width="107" height="42" border="0" id="onprom"></a><a href="javascript:;" onMouseOver="MM_swapImage('onsolution','','../Image/Solutions.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Solutions_on.gif" name="onsolution" width="109" height="42" border="0" id="onsolution"></a></div></td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td width="52"><div align="center"><font color="#FF0000">Menu</font></div></td>
    <td width="72" bordercolor="#0000FF"> <div align="center"><font color="#FF0000">Created 
        By</font></div></td>
    <td width="600"> <div align="center">:: Riski Hamdani Nasution ::</div></td>
    <td width="1">&nbsp;</td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td>Activity </td>
    <td bordercolor="#0000FF"> <div align="center"><a href="musa.asp">Musawarman</a></div></td>
    <td rowspan="9"> <table width="600" border="1" align="center" cellspacing="0" bgcolor="#003333">
        <tr> 
          <td height="73" colspan="3">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="3">&nbsp;</td>
        </tr>
        <tr bgcolor="#CCCCCC"> 
          <td colspan="3"><div align="center" class="style5">date : 
              <% =date %>
            </div></td>
        </tr>
        <tr> 
          <td width="225" rowspan="8" bgcolor="#990066"><img src="danie.gif" width="225" height="150"></td>
          <td width="180" bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Nama 
              Lengkap </font></div></td>
          <td width="331" bgcolor="#663366"><font color="#FFFFFF">Riski Hamdani 
            Nasution </font></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Panggilan 
              </font></div></td>
          <td bgcolor="#663366"><font color="#FFFFFF">Danie</font></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Tempat,Tanggal 
              Lahir</font></div></td>
          <td bgcolor="#663366"><font color="#FFFFFF">Sipirok, 12 Agustus 1983</font></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Golongan 
              Darah </font></div></td>
          <td bgcolor="#663366"><font color="#FFFFFF">O</font></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Agama 
              </font></div></td>
          <td bgcolor="#663366"><font color="#FFFFFF">Islam </font></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">No Telepon 
              </font></div></td>
          <td bgcolor="#663366"><p><font color="#FFFFFF">081317838553</font></p></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Warga 
              Negara </font></div></td>
          <td bgcolor="#663366"><font color="#FFFFFF">Indonesia</font></td>
        </tr>
        <tr> 
          <td bgcolor="#6699FF"><div align="right"><font color="#FFFF00">Email 
              </font></div></td>
          <td bgcolor="#663366"><p><font color="#FFFFFF">de_mouse83@yahoo.com 
              </font></p>
            <p><font color="#FFFFFF">danie_xxcom@yahoo.com </font></p></td>
        </tr>
        <tr bgcolor="#CCCCCC"> 
          <td colspan="3">&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>System Manager </td>
    <td bordercolor="#0000FF"> <div align="center"><a href="Riza.asp">Riza MN</a></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>Report</td>
    <td bordercolor="#0000FF"> <div align="center"><a href="Danie.asp">Riski HN</a></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td colspan="2"><div align="center">Special Thanks To </div></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td colspan="2"><p align="center">Nikko Priambodo </p>
      <p align="center">Dudi Sukma Priyadi </p>
      <p align="center">&amp; </p>
      <p align="center">TEAM</p></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td bordercolor="#0000FF">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td>&nbsp;</td>
    <td bordercolor="#0000FF">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td bordercolor="#0000FF">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td bordercolor="#0000FF"> <div align="right"><span class="style7">Date:</span> 
        <%=date%></div></td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
