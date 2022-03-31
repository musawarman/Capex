<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../../Connections/CapexConn.asp" -->
<%
Dim rscapexHD__MMColParam
rscapexHD__MMColParam = "1"
If (Request.Form("NoCapex") <> "") Then 
  rscapexHD__MMColParam = Request.Form("NoCapex")
End If
%>
<%
Dim rscapexHD
Dim rscapexHD_numRows

Set rscapexHD = Server.CreateObject("ADODB.Recordset")
rscapexHD.ActiveConnection = MM_CapexConn_STRING
rscapexHD.Source = "SELECT * FROM dbo.CapexHD WHERE NoCapex = '" + Replace(rscapexHD__MMColParam, "'", "''") + "' ORDER BY NoID ASC"
rscapexHD.CursorType = 0
rscapexHD.CursorLocation = 2
rscapexHD.LockType = 1
rscapexHD.Open()

rscapexHD_numRows = 0
%>
<%
Dim rsCapexAppr
Dim rsCapexAppr_numRows

Set rsCapexAppr = Server.CreateObject("ADODB.Recordset")
rsCapexAppr.ActiveConnection = MM_CapexConn_STRING
rsCapexAppr.Source = "SELECT * FROM dbo.CapexAppr"
rsCapexAppr.CursorType = 0
rsCapexAppr.CursorLocation = 2
rsCapexAppr.LockType = 1
rsCapexAppr.Open()

rsCapexAppr_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rscapexHD_numRows = rscapexHD_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Search No Capex ::</title>
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
//-->
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../../css/style.css" rel="stylesheet" type="text/css">
<meta http-equiv="refresh" content="5;URL=SearchNoCapex.asp">
<meta http-equiv="refresh" content="5">
<meta name="keywords" content="tes">
</head>

<body onLoad="MM_displayStatusMsg('');return document.MM_returnValue">
<div id="Layer1" style="position:absolute; left:6px; top:3px; width:159px; height:156px; z-index:1"> 
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="970" height="153" align="top">
    <param name="movie" value="../../Animasi/baner.swf">
    <param name="quality" value="high">
    <param name="SCALE" value="exactfit">
    <embed src="../../Animasi/baner.swf" width="970" height="153" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
</div>
<div align="center"> 
  <p align="left">&nbsp;</p>
  <p align="left">&nbsp;</p>
  <p align="left">&nbsp;</p>
  <p align="left">&nbsp; </p>
  <p> 
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="243" height="45">
      <param name="movie" value="../Animasi/activity.swf">
      <param name="quality" value="high">
      <embed src="../Animasi/activity.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="243" height="45"></embed> 
    </object>
  </p>
  <table width="100%" border="0" align="center">
    <tr> 
      <td width="22">&nbsp;</td>
      <td width="938"><div align="center"> 
          <h2><font color="#6699FF">.:: List Capex ::. </font></h2>
        </div></td>
      <td width="18">&nbsp;</td>
    </tr>
    <tr> 
      <td height="14" colspan="3"> <p>&nbsp; </p></td>
    </tr>
  </table>
  <div align="left"> 
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="50" height="37">
      <param name="movie" value="../../Animasi/button27.swf">
      <param name="quality" value="high">
      <embed src="../../Animasi/button27.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="50" height="37"></embed></object>
    <span class="style3">Back to Activity </span> 
    <form action="SearchNoCapex.asp" method="post" name="capexHD" id="capexHD">
      <table width="390" border="1">
        <tr bgcolor="#CCCCCC"> 
          <td width="188"><div align="center">Search</div></td>
          <td width="144"> <input name="NoCapex" type="text" id="NoCapex2">
          </td>
          <td width="36"><input name="GO" type="submit" id="GO" value="GO"></td>
        </tr>
      </table>
    </form>
    
  </div>
  <table width="970" border="1">
    
    
    <tr>
      <td width="248" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  
  
  <% If Not rscapexHD.EOF Or Not rscapexHD.BOF Then %>
  <table width="970" height="129" border="1">
    <tr bgcolor="#FF9900"> 
      <td width="120"> <div align="center">List Capex</div></td>
      <td width="893"><div align="center"><font color="#FFFFFF">Search Status 
          : Succes</font></div></td>
    </tr>
    <tr bgcolor="#663366"> 
      <td height="103" colspan="2">&nbsp; <table width="762" height="44" border="1" align="center" bordercolor="#00FFFF">
          <tr bgcolor="#6666FF"> 
            <td width="18"><div align="center"><font color="#FFFFFF">No </font></div></td>
            <td width="131"> <div align="center"><font color="#FFFFFF">NoCapex</font></div></td>
            <td width="123"> <div align="center"><font color="#FFFFFF">DivisiID</font></div></td>
            <td width="134"> <div align="center"><font color="#FFFFFF">TglCapex</font></div></td>
            <td width="152"> <div align="center"><font color="#FFFFFF">StatusCapex</font></div></td>
            <td width="164"><div align="center"><font color="#FFFFFF">Approve 
                Capex</font></div></td>
          </tr>
          <% While ((Repeat1__numRows <> 0) AND (NOT rscapexHD.EOF)) %>
          <tr bgcolor="#CCCCCC"> 
            <td height="22"> 
              <div align="center"><%=(Repeat1__index + 1)%></div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("NoCapex").Value)%> </div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("DivisiID").Value)%></div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("TglCapex").Value)%></div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("StatusCapex").Value)%></div></td>
            <td><div align="center"><font color="#FF0000"><strong><A HREF="../../MainMenu/ApprovalCapex.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "NoCapex=" & rscapexHD.Fields.Item("NoCapex").Value %>"><%=(rscapexHD.Fields.Item("ProsesApproval").Value)%></A></strong> </font> </div></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rscapexHD.MoveNext()
Wend
%>
        </table>
        <p>&nbsp; </p></td>
    </tr>
  </table>
  <% End If ' end Not rscapexHD.EOF Or NOT rscapexHD.BOF %>
  <% If rscapexHD.EOF And rscapexHD.BOF Then %>
  <table width="970" border="1" bgcolor="#663366">
    <tr> 
      <td><div align="center"><font color="#FFFFFF">Search Status</font></div></td>
    </tr>
    <tr> 
      <td><font color="#FFFF00">Silakan masukan No Capex yang valid,lihat Create 
        Capex --&gt; [ list Capex ] untuk melihat no capex yang valid.</font></td>
    </tr>
  </table>
  <br>
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" name="back" width="105" height="33" align="left" id="back">
    <param name="BGCOLOR" value="">
    <param name="movie" value="../../Animasi/button16.swf">
    <param name="quality" value="high">
    <embed src="../../Animasi/button16.swf" width="105" height="33" align="left" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" name="back" ></embed> 
  </object>
  <br>
  <br>
  <table width="970" border="1">
    <tr> 
      <td width="248" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <% End If ' end rscapexHD.EOF And rscapexHD.BOF %>
  <table width="970" border="1" bgcolor="#CC6600">
    <tr> 
      <td height="67"> <p align="center" class="style1"><a href="../../MainMenu/createcapex.asp"><font color="#FFCC00">Create 
          Capex</font></a><font color="#FFCC00"> </font><font color="#FFFFFF">..::..</font> 
          <a href="../../MainMenu/Infocapexappr.asp"><font color="#FFCC00">Proses 
          Approval Capex</font></a> <font color="#FFFFFF"> ..::..</font> <a href="../../MainMenu/ListCapexApproved.asp"><font color="#FFCC00"> 
          Create AOC</font></a><font color="#FFFFFF"> ..::..</font> <a href="../../MainMenu/InfoAOCappr.asp"><font color="#FFCC00"> 
          Create Approval AOC</font></a><font color="#FFFFFF"> ..::.. </font><a href="../../MainMenu/hold%20Capex.asp"><font color="#FFCC00"> 
          Hold Capex</font></a> <font color="#FFFFFF">..::..</font> <a href="../../MainMenu/hold%20AOC.asp"><font color="#FFCC00">Hold 
          AOC</font></a> </p>
        <p align="center" class="style1"><a href="../../MainMenu/SearchNoAOCforPayment.asp"><font color="#FFCC00">Payman 
          Estimation</font></a> <font color="#FFFFFF">..::.. </font><a href="../../MainMenu/Payman%20actual.asp"><font color="#FFCC00">Payman 
          Actual</font></a> </p></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rscapexHD.Close()
Set rscapexHD = Nothing
%>
<%
rsCapexAppr.Close()
Set rsCapexAppr = Nothing
%>
