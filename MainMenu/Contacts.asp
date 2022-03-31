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

<body background="../Image/bgact.gif" onLoad="MM_preloadImages('../Image/homee.gif','../Image/aboutus.gif','../Image/contact.gif','../Image/Promotion.gif','../Image/Solutions.gif')">
<table width="740" height="1400" border="1" align="center" bordercolor="#0000FF">
  <tr> 
    <td height="36" colspan="5"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="730" height="100" align="top">
        <param name="movie" value="../Animasi/baner.swf">
        <param name="quality" value="high">
        <param name="SCALE" value="exactfit">
        <embed src="../Animasi/baner.swf" width="730" height="100" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object></td>
  </tr>
  <tr> 
    <td width="2" height="44" bordercolor="#3366FF">&nbsp;</td>
    <td width="2" bordercolor="#3366FF">&nbsp;</td>
    <td width="219" bordercolor="#3366FF"> <div align="center">Welcome <%= Session("UpdateUsr") %></div></td>
    <td colspan="2" bordercolor="#3366FF"> <div align="center"><a href="mainpage.asp" onMouseOver="MM_swapImage('onhome','','../Image/homee.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/homee_on.gif" name="onhome" width="83" height="42" border="0" id="onhome"></a><a href="../ABOUT/about.asp" onMouseOver="MM_swapImage('onabout','','../Image/aboutus.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/aboutus_on.gif" name="onabout" width="95" height="42" border="0" id="onabout"></a><a href="Contacts.asp" onMouseOver="MM_swapImage('oncontact','','../Image/contact.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/contact_on.gif" name="oncontact" width="95" height="42" border="0" id="oncontact"></a><a href="Promotion.asp" onMouseOver="MM_swapImage('onprom','','../Image/Promotion.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Promotion_on.gif" name="onprom" width="107" height="42" border="0" id="onprom"></a><a href="javascript:;" onMouseOver="MM_swapImage('onsolution','','../Image/Solutions.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Solutions_on.gif" name="onsolution" width="109" height="42" border="0" id="onsolution"></a></div></td>
  </tr>
  <tr> 
    <td height="23">&nbsp;</td>
    <td>&nbsp;</td>
    <td><div align="center">:: Contacts Information ::</div></td>
    <td width="490"> <div align="center">Silahkan hubungi kami untuk informasi 
        .. </div></td>
    <td width="1">&nbsp;</td>
  </tr>
  <tr> 
    <td height="34">&nbsp;</td>
    <td>&nbsp;</td>
    <td bordercolor="0"><b>Kantor Pusat </b><br>
      Jl. Kemang Raya No. 67 <br>
      Jakarta 12730<br>
      Tel. (021) 7193888 <br>
      Fax. (021) 7193889 <br> <br> <b>Divisi Breeding</b><br>
      Jl. Raya Parung KM 20 Parung<br>
      Bogor - Jawa Barat<br>
      Tel. (0251) 611143<br>
      Fax. (0251) 611655<br> <br> <br> <b>Divisi Contract Farming</b><br>
      Bojong Sari No. 26 Sawangan,<br>
      Bogor-Jawa Barat<br>
      Tel. (0251) 619789<br>
      Fax. (0251) 619790<br> <br>
      Jl. Raya Sidoarjo - Krian Desa Ketimang,<br>
      Wonoayu-Sidoarjo, Jawa Timur<br>
      Tel. (031) 8852804 - 06<br>
      Fax. (031) 8852810<br> <br>
      Jl. Patimura KM 1 Canden<br>
      Salatiga 50711 Jawa Tengah<br>
      Tel. (0298) 327353 <br>
      Fax.(0298) 313492<br> <br> <b>Divisi Feed Mill<br>
      </b>Jl. Raya Serang KM 13 Balaraja,<br>
      Tangerang - Banten <br>
      Tel. (021) 5953888 <br>
      Fax.(021) 5950150 <br> <br>
      Jl. Raya Sidoarjo - Krian Desa Ketimang, <br>
      Wonoayu-Sidoarjo, Jawa Timur <br>
      Tel. (031) 8852804 - 06 <br>
      Fax. (031) 8852810 <br> <br>
      Jl. Ir. Sutami KM 12 Desa Suka Negara Tanjung Bintang <br>
      Lampung Selatan <br>
      Tel. (0721) 351175 <br>
      Fax. (0721) 351173<br> <br> <b>Divisi Slaughterhouse<br>
      </b>Jl. Raya Parung KM 19 Parung, <br>
      Bogor-Jawa Barat <br>
      Tel. (0251) 611862 <br>
      Fax. (0251) 611079<br> <br>
      Jl. Patimura KM 1 Canden <br>
      Salatiga 50711 Jawa Tengah <br>
      Tel. (0298) 327353 <br>
      Fax. (0298) 313492<br> <br> <b>PT Sierad Industries</b><br>
      Poultry Equipment<br>
      Jl. Modern Industri No. 24 <br>
      Kawasan Industri Modern Cikande - Serang, Banten <br>
      Tel. (0254) 402536 <br>
      Fax. (0254) 402538<br> <br> <b>PT Sierad Biotek</b><br>
      Animal Pharmaceutical &amp; Health Care Products<br>
      Jl. M.H. Thamrin Blok A-10 No. 3 <br>
      Lippo Cikarang, Bekasi 17550 - Jawa Barat <br>
      Tel. (021) 8972264 - 66 <br>
      Fax. (021) 8972268 <br> <br> <b>PT Dwipamina Nusantara</b> <br>
      Fishmeal Product<br>
      Dusun Ketapang Desa Pengambengan Negara - Jembrana, Bali <br>
      Tel. (0365) 42147 <br>
      Fax.(0365) 42148 <br> <br> <b>PT Sierad Pangan</b> <br>
      Hartz Restaurant<br>
      Jl. Raya Kebayoran Lama No. 220 <br>
      Jakarta Selatan 12220 <br>
      Tel. (021) 5360482 <br>
      Fax. (021) 5357932<br> <br> <b>PT Wendy Citarasa</b> <br>
      Wendy's Fast Foods<br>
      Jl. Raya Kebayoran Lama No. 220 <br>
      Jakarta Selatan 12220 <br>
      Tel. (021) 5350481 <br>
      Fax. (021) 5357932<br> <br> <br> <b></b></td>
    <td rowspan="2">PT Sierad Produce Tbk adalah gabungan dari 4 perusahaan pada 
      tahun 2001 yang bergerak di satu bidang bisnis utama di bawah naungan Sierad 
      Group. Empat perusahaan ini adalah PT Anwar Sierad Tbk, PT Sierad Produce 
      Tbk, PT Sierad Feedmill dan PT Sierad Grains.<br> <br> &nbsp;&nbsp;&nbsp;&nbsp;Sierad 
      Produce, dahulu bernama PT. Betara Darma ekspor impor, berdiri pada tanggal 
      6 September 1985. Nama Sierad mulai digunakan pada tanggal 27 Desember 1996 
      saat persiapan untuk public listing yang cukup berhasil di Jakarta Stock 
      Exchange. Bisnis utama perusahaan ini meliputi produksi pakan ternak olahan, 
      breeding, produksi anak ayam, kemitraan, rumah potong ayam dan pembuatan 
      produk olahan bernilai tambah lainnya.</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="2">&nbsp;</td>
    <td bordercolor="0"><div align="center">[ <a href="<%= MM_Logout %>">Logout</a> 
        ] </div></td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
