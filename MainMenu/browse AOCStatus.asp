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
Dim rsAOCHD
Dim rsAOCHD_numRows

Set rsAOCHD = Server.CreateObject("ADODB.Recordset")
rsAOCHD.ActiveConnection = MM_CapexConn_STRING
rsAOCHD.Source = "SELECT NoAOC, NoCapex, VendorID, StatusAOC, TglAOC, Hold_AOC_By  FROM dbo.AOCHD  ORDER BY UpdateDate ASC"
rsAOCHD.CursorType = 0
rsAOCHD.CursorLocation = 2
rsAOCHD.LockType = 1
rsAOCHD.Open()

rsAOCHD_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsAOCHD_numRows = rsAOCHD_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsAOCHD_total
Dim rsAOCHD_first
Dim rsAOCHD_last

' set the record count
rsAOCHD_total = rsAOCHD.RecordCount

' set the number of rows displayed on this page
If (rsAOCHD_numRows < 0) Then
  rsAOCHD_numRows = rsAOCHD_total
Elseif (rsAOCHD_numRows = 0) Then
  rsAOCHD_numRows = 1
End If

' set the first and last displayed record
rsAOCHD_first = 1
rsAOCHD_last  = rsAOCHD_first + rsAOCHD_numRows - 1

' if we have the correct record count, check the other stats
If (rsAOCHD_total <> -1) Then
  If (rsAOCHD_first > rsAOCHD_total) Then
    rsAOCHD_first = rsAOCHD_total
  End If
  If (rsAOCHD_last > rsAOCHD_total) Then
    rsAOCHD_last = rsAOCHD_total
  End If
  If (rsAOCHD_numRows > rsAOCHD_total) Then
    rsAOCHD_numRows = rsAOCHD_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = rsAOCHD
MM_rsCount   = rsAOCHD_total
MM_size      = rsAOCHD_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsAOCHD_first = MM_offset + 1
rsAOCHD_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsAOCHD_first > MM_rsCount) Then
    rsAOCHD_first = MM_rsCount
  End If
  If (rsAOCHD_last > MM_rsCount) Then
    rsAOCHD_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = MM_keepMove & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
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
  <div align="left"> 
    <div align="center"> 
      <div align="center"> 
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
                      Kami</font></a></font><font color="#FF0000">&nbsp; </font>| 
                      <a href="../karir.asp"><font color="#006600">Karir </font></a>| 
                      <a href="../link.asp"><font color="#006600">Links </font></a>| 
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
                  <td width="296"><div align="right"><font color="#FFFFFF">Time 
                      : <strong><%=time()%></strong></font></div></td>
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
                      <strong>STATUS AOC</strong></font></div></td>
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
                      <table width="700" border="1" align="center">
                        <tr bgcolor="#6699FF"> 
                          <td> <div align="center"><font color="#FFFFFF">NoAOC</font></div></td>
                          <td> <div align="center"><font color="#FFFFFF">NoCapex</font></div></td>
                          <td> <div align="center"><font color="#FFFFFF">VendorID</font></div></td>
                          <td> <div align="center"><font color="#FFFFFF">StatusAOC</font></div></td>
                          <td> <div align="center"><font color="#FFFFFF">TglAOC</font></div></td>
                          <td> <div align="center"><font color="#FFFFFF">Hold_AOC_By</font></div></td>
                        </tr>
                        <% While ((Repeat1__numRows <> 0) AND (NOT rsAOCHD.EOF)) %>
                        <tr bgcolor="#CCCCCC"> 
                          <td height="16"> <div align="center"><%=(rsAOCHD.Fields.Item("NoAOC").Value)%></div></td>
                          <td> <div align="center"><%=(rsAOCHD.Fields.Item("NoCapex").Value)%></div></td>
                          <td> <div align="center"><%=(rsAOCHD.Fields.Item("VendorID").Value)%></div></td>
                          <td> <div align="center"><%=(rsAOCHD.Fields.Item("StatusAOC").Value)%></div></td>
                          <td> <div align="center"><%=(rsAOCHD.Fields.Item("TglAOC").Value)%></div></td>
                          <td> <div align="center"><%=(rsAOCHD.Fields.Item("Hold_AOC_By").Value)%></div></td>
                        </tr>
                        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsAOCHD.MoveNext()
Wend
%>
                      </table>
                      <table border="0" width="740" align="center">
                        <tr> 
                          <td width="42%" align="center"> 
                            <% If MM_offset <> 0 Then %>
                            <a href="<%=MM_moveFirst%>"><img src="../Image/first.gif" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; first page');return document.MM_returnValue"></a> 
                            <% End If ' end MM_offset <> 0 %>
                          </td>
                          <td width="8%" align="center"> 
                            <% If MM_offset <> 0 Then %>
                            <a href="<%=MM_movePrev%>"><img src="../Image/previous.gif" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; previous page');return document.MM_returnValue"></a> 
                            <% End If ' end MM_offset <> 0 %>
                          </td>
                          <td width="8%" align="center"> 
                            <% If Not MM_atTotal Then %>
                            <a href="<%=MM_moveNext%>"><img src="../Image/next.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt;next page');return document.MM_returnValue"></a> 
                            <% End If ' end Not MM_atTotal %>
                          </td>
                          <td width="42%" align="center"> 
                            <% If Not MM_atTotal Then %>
                            <a href="<%=MM_moveLast%>"><img src="../Image/last.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; last page');return document.MM_returnValue"></a> 
                            <% End If ' end Not MM_atTotal %>
                          </td>
                        </tr>
                      </table>
                      <p><img src="../../img/garis1.gif" width="400" height="1"></p>
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
                              <a href="Sman.asp"><font color="#006600">SYSTEM 
                              MANAGER</font></a><font color="#006600"> </font>| 
                              <a href="Activity.asp"><font color="#006600">ACTIVITIES</font></a><font color="#006600"> 
                              </font>|<a href="../Reports/reportListing.asp"> 
                              <font color="#006600">REPORT</font></a><font color="#006600"> 
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
      </div>
    </div>
    
  </div>
</div>
</body>
</html>
<%
rsAOCHD.Close()
Set rsAOCHD = Nothing
%>
