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
<!--#include file="../Connections/CapexConn.asp" -->
<%
Dim rscapexdt
Dim rscapexdt_numRows

Set rscapexdt = Server.CreateObject("ADODB.Recordset")
rscapexdt.ActiveConnection = MM_CapexConn_STRING
rscapexdt.Source = "SELECT * FROM dbo.CapexDT ORDER BY UpdateDate ASC"
rscapexdt.CursorType = 0
rscapexdt.CursorLocation = 2
rscapexdt.LockType = 1
rscapexdt.Open()

rscapexdt_numRows = 0
%>
<%
Dim rscapexhd
Dim rscapexhd_numRows

Set rscapexhd = Server.CreateObject("ADODB.Recordset")
rscapexhd.ActiveConnection = MM_CapexConn_STRING
rscapexhd.Source = "SELECT * FROM dbo.CapexHD"
rscapexhd.CursorType = 0
rscapexhd.CursorLocation = 2
rscapexhd.LockType = 1
rscapexhd.Open()

rscapexhd_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rscapexdt_numRows = rscapexdt_numRows + Repeat1__numRows
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Capex Detail ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-size: 18px}
.style2 {font-size: 36px}
.style3 {font-size: 16px}
body,td,th {
	color: #0000FF;
}
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

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}

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
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="../Image/bgact.gif" onLoad="MM_displayStatusMsg('Create Capex');return document.MM_returnValue;MM_preloadImages('../Image/CreateCapex_on.gif','../Image/ApprovalCapex_on.gif','../Image/CreateAOC_on.gif','../Image/Approvalaoc_on.gif','../Image/holdcapex_on.gif','../Image/holdAoc_on.gif','../Image/Estimation_on.gif','../Image/Actual_on.gif')">
<div align="center"> 
  <div align="center"> 
    <div align="center"> 
      <table width="800" border="1" align="center" bordercolor="#006600">
        <tr bordercolor="#006600"> 
          <td colspan="2"> <div align="left"><img src="../Image/sieradonline.gif" width="222" height="85"> 
            </div>
            <div align="right"><font color="#006600">Date : 
              <script name="current" src="../../GeneratedItems/current.js" language="JavaScript1.2"></script>
              </font></div></td>
        </tr>
        <tr bordercolor="#006600" bgcolor="#CCCCCC"> 
          <td> <div align="left"><font color="#006600">Welcome <%= Session("UpdateUsr") %></font></div></td>
          <td width="300"> <div align="center"><font color="#009900"><a href="../contact.asp"><font color="#006600">Hubungi 
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
      <table width="800" border="1" align="center" bordercolor="#CCCCCC" bgcolor="#006600">
        <tr> 
          <td width="150" height="23"><div align="center"><a href="createcapex.asp" onMouseOver="MM_swapImage('Image1','','../Image/CreateCapex_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/CreateCapex.gif" name="Image1" width="150" height="23" border="0" id="Image1"></a></div></td>
          <td width="330" rowspan="8" bgcolor="#006600"><div align="center"> 
              <p> 
                <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="323" height="122" align="top">
                  <param name="movie" value="../Animasi/anakayam.swf">
                  <param name="quality" value="high">
                  <param name="SCALE" value="exactfit">
                  <embed src="../Animasi/anakayam.swf" width="323" height="122" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
              </p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp; </p>
            </div></td>
          <td width="296"><div align="right"><font color="#FFFFFF">Time : <strong><%=time()%></strong></font></div></td>
        </tr>
        <tr> 
          <td><div align="center"><a href="Infocapexappr.asp" onMouseOver="MM_swapImage('Image2','','../Image/ApprovalCapex_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/ApprovalCapex.gif" name="Image2" width="150" height="23" border="0" id="Image2"></a></div></td>
          <td rowspan="7"><div align="left"> 
              <p>&nbsp;</p>
              <p>&nbsp;</p>
            </div></td>
        </tr>
        <tr> 
          <td><div align="center"><a href="ListCapexApproved.asp" onMouseOver="MM_swapImage('Image3','','../Image/CreateAOC_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/CreateAOC.gif" name="Image3" width="150" height="23" border="0" id="Image3"></a></div></td>
        </tr>
        <tr> 
          <td><div align="center"><a href="InfoAOCappr.asp" onMouseOver="MM_swapImage('Image4','','../Image/Approvalaoc_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Approvalaoc.gif" name="Image4" width="150" height="23" border="0" id="Image4"></a></div></td>
        </tr>
        <tr> 
          <td height="16"> <div align="center"><a href="hold%20capex.asp" onMouseOver="MM_swapImage('Image5','','../Image/holdcapex_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/holdcapex.gif" name="Image5" width="150" height="23" border="0" id="Image5"></a></div></td>
        </tr>
        <tr> 
          <td height="24"><a href="hold%20AOC.asp" onMouseOver="MM_swapImage('Image6','','../Image/holdAoc_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/holdAoc.gif" name="Image6" width="150" height="23" border="0" id="Image6"></a></td>
        </tr>
        <tr> 
          <td height="24"><a href="SearchNoAOCforPayment.asp" onMouseOver="MM_swapImage('Image7','','../Image/Estimation_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Estimation.gif" name="Image7" width="150" height="23" border="0" id="Image7"></a></td>
        </tr>
        <tr> 
          <td height="27"> <div align="center"><a href="Payman%20actual.asp" onMouseOver="MM_swapImage('Image8','','../Image/Actual_on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../Image/Actual.gif" name="Image8" width="150" height="23" border="0" id="Image8"></a></div></td>
        </tr>
      </table>
      <table width="800" border="0" align="center" bordercolor="#FF6600" bgcolor="#006600">
        <tr> 
          <td><div align="center"><img src="../../BREEDING/img/spacer.gif" width="795" height="10"></div></td>
        </tr>
      </table>
      <table width="806" border="0" align="center" background="../../img/bg.gif">
        <tr> 
          <td width="800" height="20"><div align="left"><font color="#FF0000" size="3" face="Courier New, Courier, mono">&gt;&gt; 
              <strong>LIST CAPEX</strong></font></div></td>
        </tr>
        <tr> 
          <td><div align="left"><img src="../Image/spacer.jpg" width="200" height="2"></div></td>
        </tr>
        <tr> 
          <td height="215"> <div align="center"> 
              <table width="800" border="0">
                <tr> 
                  <td height="3"> <div align="center"><img src="../../img/garis1.gif" width="600" height="1"></div></td>
                </tr>
              </table>
              <p><img src="../../img/garis1.gif" width="400" height="1"></p>
              <table width="785" border="1">
                <tr bgcolor="#6666FF"> 
                  <td width="18" bordercolor="#00FFFF"><div align="center"><font color="#FFFFFF">No 
                      </font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">NoCapex</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Seq</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">BudgetID</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">NamaBarang</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Qty</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Price</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">QtyPrice</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">IDRPRice</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">IDRQTYPrice</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">CurrencyID</font></div></td>
                  <td> <div align="center"><font color="#FFFFFF">Post By</font></div></td>
                </tr>
                <% While ((Repeat1__numRows <> 0) AND (NOT rscapexdt.EOF)) %>
                <tr bgcolor="#CCCCCC"> 
                  <td bordercolor="#00FFFF"> <div align="center"><%=(Repeat1__index + 1)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("NoCapex").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("Seq").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("BudgetID").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("NamaBarang").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("Qty").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("Price").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("QtyPrice").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("IDRPRice").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("IDRQTYPrice").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("CurrencyID").Value)%></div></td>
                  <td> <div align="center"><%=(rscapexdt.Fields.Item("UpdateUsr").Value)%></div></td>
                </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rscapexdt.MoveNext()
Wend
%>
              </table>
              <form name="form2" method="post" action="">
                <input name="Submit2" type=button class=btn onclick=history.back() value="Back">
              </form>
              <p><img src="../../img/garis1.gif" width="400" height="1"></p>
              <p>&nbsp;</p>
            </div></td>
        </tr>
        <tr> 
          <td><div align="center"><img src="../../img/garis1.gif" width="600" height="1"></div></td>
        </tr>
        <tr> 
          <td height="47"><div align="center"> 
              <p>&nbsp;</p>
              <table width="800" border="1" bordercolor="#FF6600">
                <tr> 
                  <td><div align="center"><font color="#009900"><font color="#006600">-- 
                      HOME</font></font><font color="#006600">&nbsp; </font>| 
                      <a href="Sman.asp"><font color="#006600">SYSTEM MANAGER</font></a><font color="#006600"> 
                      </font>| <a href="Activity.asp"><font color="#006600">ACTIVITIES</font></a><font color="#006600"> 
                      </font>|<a href="../Reports/reportListing.asp"> <font color="#006600">REPORT</font></a><font color="#006600"> 
                      | FAQ --</font></div></td>
                </tr>
              </table>
              <p><img src="../Image/bannerrg.jpg" width="800" height="20"></p>
            </div></td>
        </tr>
      </table>
    </div>
  </div>
</div>
</body>
</html>
<%
rscapexdt.Close()
Set rscapexdt = Nothing
%>
<%
rscapexhd.Close()
Set rscapexhd = Nothing
%>
