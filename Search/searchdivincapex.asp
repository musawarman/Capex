<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="../Connections/CapexConn.asp" -->
<%
Dim rscapexHD__MMColParam
rscapexHD__MMColParam = "1"
If (Request.Form("DivisiID") <> "") Then 
  rscapexHD__MMColParam = Request.Form("DivisiID")
End If
%>
<%
Dim rscapexHD
Dim rscapexHD_numRows

Set rscapexHD = Server.CreateObject("ADODB.Recordset")
rscapexHD.ActiveConnection = MM_CapexConn_STRING
rscapexHD.Source = "SELECT * FROM dbo.CapexHD WHERE DivisiID = '" + Replace(rscapexHD__MMColParam, "'", "''") + "' ORDER BY NoID ASC"
rscapexHD.CursorType = 0
rscapexHD.CursorLocation = 2
rscapexHD.LockType = 1
rscapexHD.Open()

rscapexHD_numRows = 0
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
<title>:: List Capex ::</title>
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
</head>

<body onLoad="MM_displayStatusMsg('Create Capex');return document.MM_returnValue">
<div id="Layer1" style="position:absolute; left:6px; top:3px; width:159px; height:156px; z-index:1"> 
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="1125" height="153" align="top">
    <param name="movie" value="../Animasi/baner.swf">
    <param name="quality" value="high">
    <param name="SCALE" value="exactfit">
    <embed src="../Animasi/baner.swf" width="1125" height="153" align="top" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
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
  <p align="left" class="style2">
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="50" height="37">
      <param name="BGCOLOR" value="">
      <param name="movie" value="../Animasi/button27.swf">
      <param name="quality" value="high">
      <embed src="../Animasi/button27.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="50" height="37" ></embed>
    </object> 
    <span class="style3">Back to Activity
  </span> </p>
  <table width="1125" height="44" border="1">
    
    
    
    
    
    
    <tr>
      <td width="248" height="38" bgcolor="#669900"> <div align="left">
          <form action="searchdivincapex.asp" method="post" name="divisiform" id="divisiform">
            <table width="325" border="1" align="left" bgcolor="#663399">
              <tr> 
                <td width="127"> <div align="center"><font color="#FFFFFF">Search 
                    For Divisi ID</font></div></td>
                <td width="144"> <input name="divisiID" type="text" id="divisi3"></td>
                <td width="40"> <input type="submit" name="Submit" value="GO"></td>
              </tr>
            </table>
          </form>
        </div></td>
    </tr>
  </table>
  
  
  
  
  
  <table width="1000" height="129" border="1">
    <tr bgcolor="#FF9900"> 
      <td width="120"> <div align="center">List Capex</div></td>
      <td width="893">&nbsp;</td>
    </tr>
    <tr bgcolor="#663366"> 
      <td height="98" colspan="2">&nbsp; <table width="666" border="1" align="center" bordercolor="#00FFFF">
          <tr bgcolor="#6666FF"> 
            <td width="18"><div align="center"><font color="#FFFFFF">No </font></div></td>
            <td width="129"> <div align="center"><font color="#FFFFFF">NoCapex</font></div></td>
            <td width="121"> <div align="center"><font color="#FFFFFF">DivisiID</font></div></td>
            <td width="132"> <div align="center"><font color="#FFFFFF">TglCapex</font></div></td>
            <td width="150"> <div align="center"><font color="#FFFFFF">StatusCapex</font></div></td>
            <td width="76"><div align="center"><font color="#FFFFFF">Detail Capex</font></div></td>
          </tr>
          <% While ((Repeat1__numRows <> 0) AND (NOT rscapexHD.EOF)) %>
          <tr bgcolor="#CCCCCC"> 
            <td> 
              <div align="center"><%=(Repeat1__index + 1)%></div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("NoCapex").Value)%> </div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("DivisiID").Value)%></div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("TglCapex").Value)%></div></td>
            <td><div align="center"><%=(rscapexHD.Fields.Item("StatusCapex").Value)%></div></td>
            <td><div align="center"><A HREF="../MainMenu/Detail%20Capex.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "NoCapex=" & rscapexHD.Fields.Item("NoCapex").Value %>"><%=(rscapexHD.Fields.Item("Detail").Value)%></A></div></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rscapexHD.MoveNext()
Wend
%>
        </table>
        <p> 
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="100" height="23" hspace="54" vspace="0" align="middle">
            <param name="movie" value="../MainMenu/Add.swf">
            <param name="quality" value="high">
            <param name="base" value=".">
            <param name="bgcolor" value="#663366">
            <embed src="../MainMenu/Add.swf" width="100" height="23" hspace="54" vspace="0" align="middle" base="."  quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" bgcolor="#663366"></embed> 
          </object>
        </p></td>
    </tr>
  </table>
  
  <p align="center" class="style1">
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" name="back" width="105" height="33" align="left" id="back">
      <param name="BGCOLOR" value="">
      <param name="movie" value="../Animasi/button16.swf">
      <param name="quality" value="high">
      <embed src="../Animasi/button16.swf" width="105" height="33" align="left" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" name="back" ></embed> 
    </object>
  </p>
  <p align="left" class="style1">&nbsp;</p>
  <table width="1125" border="1">
    <tr>
      <td bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <table width="1125" border="1" bgcolor="#CC6600">
    <tr> 
      <td height="67"> <p align="center" class="style1"><a href="../MainMenu/createcapex.asp"><font color="#FFCC00">Create 
          Capex</font></a><font color="#FFCC00"> </font><font color="#FFFFFF">..::..</font> 
          <a href="../MainMenu/Infocapexappr.asp"><font color="#FFCC00">Proses 
          Approval Capex</font></a> <font color="#FFFFFF"> ..::..</font> <a href="../MainMenu/create%20AOC.asp"><font color="#FFCC00"> 
          Create AOC</font></a><font color="#FFFFFF"> ..::..</font> <a href="../MainMenu/create%20approval.asp"><font color="#FFCC00"> 
          Create Approval </font></a><font color="#FFFFFF"> ..::.. </font><a href="../MainMenu/hold%20capex.asp"><font color="#FFCC00"> 
          Hold Capex</font></a> <font color="#FFFFFF">..::..</font> <a href="../MainMenu/hold%20AOC.asp"><font color="#FFCC00">Hold 
          AOC</font></a> </p>
        <p align="center" class="style1"><a href="../MainMenu/Payman%20estimation.asp"><font color="#FFCC00">Payman 
          Estimation</font></a> <font color="#FFFFFF">..::.. </font><a href="../MainMenu/Payman%20actual.asp"><font color="#FFCC00">Payman 
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
