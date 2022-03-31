<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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

<body background="../Image/bgact.gif" onLoad="MM_preloadImages('../Image/homee.gif','../Image/aboutus.gif','../Image/contact.gif','../Image/Promotion.gif','../Image/Solutions.gif')">
<table width="740" height="527" border="1" align="center" bordercolor="#0000FF">
  <tr> 
    <td height="36" colspan="5"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="730" height="100" align="top">
        <param name="movie" value="../Animasi/baner.swf">
        <param name="quality" value="high">
        <param name="SCALE" value="exactfit">
        <embed src="../Animasi/baner.swf" width="730" height="100" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object></td>
  </tr>
  <tr> 
    <td width="3" height="44" bordercolor="#3366FF">&nbsp;</td>
    <td width="3" bordercolor="#3366FF">&nbsp;</td>
    <td width="117" bordercolor="#3366FF"> <div align="left">Welcome <%= Session("UpdateUsr") %></div></td>
    <td colspan="2" bordercolor="#3366FF"> <div align="center"><a href="mainpage.asp" onMouseOver="MM_swapImage('onhome','','../Image/homee.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/homee_on.gif" name="onhome" width="83" height="42" border="0" id="onhome"></a><a href="../ABOUT/about.asp" onMouseOver="MM_swapImage('onabout','','../Image/aboutus.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/aboutus_on.gif" name="onabout" width="95" height="42" border="0" id="onabout"></a><a href="Contacts.asp" onMouseOver="MM_swapImage('oncontact','','../Image/contact.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/contact_on.gif" name="oncontact" width="95" height="42" border="0" id="oncontact"></a><a href="Promotion.asp" onMouseOver="MM_swapImage('onprom','','../Image/Promotion.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Promotion_on.gif" name="onprom" width="107" height="42" border="0" id="onprom"></a><a href="javascript:;" onMouseOver="MM_swapImage('onsolution','','../Image/Solutions.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Solutions_on.gif" name="onsolution" width="109" height="42" border="0" id="onsolution"></a></div></td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="center"></div></td>
    <td width="582"> <div align="center">:: Promotion ::</div></td>
    <td width="1">&nbsp;</td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td>&nbsp;</td>
    <td rowspan="9">&nbsp;</td>
    <td rowspan="9">PT.Sierad Produce Tbk. Dalam mengupayakan pengembangan lingkungan 
      yang aman dan kondusif, maka perusahaan melakukan kegiatan sosial dan memberi 
      bantuan seperti:<br>
      - Melaksanakan kegiatan perayaan keagamaan untuk karyawan dan lingkungan 
      setempat.<br>
      - Melakukan kegiatan Sierad Peduli dengan membagi telur rebus kepada anak 
      murid SD di sekitar perusahaan agar meningkatkan gizi anak.<br>
      - Memberikan kambing bergulir kepada warga yang membutuhkan usaha ternak 
      kambing sehingga dapat menambah pendapat keluarga.<br>
      Menjalin tali silaturahim antar perusahaan dengan para tokoh masyarakat 
      dan aparat desa serta Muspika setempat. <br></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
