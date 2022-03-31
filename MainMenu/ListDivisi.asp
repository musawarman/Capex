
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
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="8"
MM_authFailedURL="login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (false Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
Dim rsDivisi
Dim rsDivisi_numRows

Set rsDivisi = Server.CreateObject("ADODB.Recordset")
rsDivisi.ActiveConnection = MM_CapexConn_STRING
rsDivisi.Source = "SELECT * FROM dbo.Divisi ORDER BY UpdateDate ASC"
rsDivisi.CursorType = 0
rsDivisi.CursorLocation = 2
rsDivisi.LockType = 1
rsDivisi.Open()

rsDivisi_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsDivisi_numRows = rsDivisi_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsDivisi_total
Dim rsDivisi_first
Dim rsDivisi_last

' set the record count
rsDivisi_total = rsDivisi.RecordCount

' set the number of rows displayed on this page
If (rsDivisi_numRows < 0) Then
  rsDivisi_numRows = rsDivisi_total
Elseif (rsDivisi_numRows = 0) Then
  rsDivisi_numRows = 1
End If

' set the first and last displayed record
rsDivisi_first = 1
rsDivisi_last  = rsDivisi_first + rsDivisi_numRows - 1

' if we have the correct record count, check the other stats
If (rsDivisi_total <> -1) Then
  If (rsDivisi_first > rsDivisi_total) Then
    rsDivisi_first = rsDivisi_total
  End If
  If (rsDivisi_last > rsDivisi_total) Then
    rsDivisi_last = rsDivisi_total
  End If
  If (rsDivisi_numRows > rsDivisi_total) Then
    rsDivisi_numRows = rsDivisi_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsDivisi_total = -1) Then

  ' count the total records by iterating through the recordset
  rsDivisi_total=0
  While (Not rsDivisi.EOF)
    rsDivisi_total = rsDivisi_total + 1
    rsDivisi.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsDivisi.CursorType > 0) Then
    rsDivisi.MoveFirst
  Else
    rsDivisi.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsDivisi_numRows < 0 Or rsDivisi_numRows > rsDivisi_total) Then
    rsDivisi_numRows = rsDivisi_total
  End If

  ' set the first and last displayed record
  rsDivisi_first = 1
  rsDivisi_last = rsDivisi_first + rsDivisi_numRows - 1
  
  If (rsDivisi_first > rsDivisi_total) Then
    rsDivisi_first = rsDivisi_total
  End If
  If (rsDivisi_last > rsDivisi_total) Then
    rsDivisi_last = rsDivisi_total
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

Set MM_rs    = rsDivisi
MM_rsCount   = rsDivisi_total
MM_size      = rsDivisi_numRows
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
rsDivisi_first = MM_offset + 1
rsDivisi_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsDivisi_first > MM_rsCount) Then
    rsDivisi_first = MM_rsCount
  End If
  If (rsDivisi_last > MM_rsCount) Then
    rsDivisi_last = MM_rsCount
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
<title>::List Divisi ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-size: 18px}
.style4 {font-size: 14px}
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
<style type="text/css">
<!--
.style5 {color: #FFFFFF}
.style7 {color: #FFFFFF; font-weight: bold; }
.style9 {
	font-family: verdana;
	font-size: 12px;
	font-weight: bold;
}
-->
</style>
<% session.timeout=20%>
</head>

<body bgcolor="#FFFFFF" background="../Image/bg.gif">
<div align="center"> 
  <table width="600" border="0">
    <tr> 
      <td colspan="2"><img src="../Image/banner2.gif" width="750" height="100"></td>
    </tr>
    <tr> 
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: List Divisi 
          ::. </font></h3></td>
    </tr>
    <tr> 
      <td width="475"><div align="left"><span class="style4"> <strong><span class="style4">
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="50" height="37" hspace="10" align="absmiddle" onMouseOver="MM_displayStatusMsg('rz : delete vendor confirmation -&gt; back to system manager');return document.MM_returnValue">
            <param name="BASE" value=".">
            <param name="BGCOLOR" value="">
            <param name="movie" value="back2sm.swf">
            <embed src="back2sm.swf" width="50" height="37" hspace="10" align="absmiddle" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" base="." ></embed> 
          </object>
          </span></strong> <span class="style4"><font color="#0000FF" size="-1" face="Arial, Helvetica, sans-serif">back 
          to system manager</font> </span> <span class="style9"></span> </span></div></td>
      <td width="120">
<div align="center"><font size="2"><a href="<%= MM_Logout %>" target="_parent">Logout</a> 
          |<strong> <a href="../Search/Search.asp" target="_parent">Search</a></strong></font></div></td>
    </tr>
    <tr bgcolor="#669900"> 
      <td colspan="2"><div align="center"><span class="style7">Welcome <%= Session("UpdateUsr") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td colspan="2"><div align="center"><font color="#0000A0"><strong>..:: Records 
          <%=(rsDivisi_first)%> to <%=(rsDivisi_last)%> of <%=(rsDivisi_total)%> ::.. </strong></font></div></td>
    </tr>
  </table>
  <br>
  <table border="1" align="center" cellspacing="0" bordercolor="#FFFFFF">
    <tr bgcolor="#6666FF"> 
      <td width="17"><div align="center" class="style5">NO</div></td>
      <td width="36"><div align="center" class="style5">Delete</div></td>
      <td width="40"><div align="center" class="style5">Update</div></td>
      <td width="97"><div align="center" class="style5">DivisiID</div></td>
      <td width="121"><div align="center" class="style5">CompanyID</div></td>
      <td width="116"><div align="center" class="style5">DivisiName</div></td>
      <td width="113"><div align="center" class="style5">Post By </div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsDivisi.EOF)) %>
    <tr bgcolor="#CCCCCC"> 
      <td><div align="center"><%=(Repeat1__index + 1)%></div></td>
      <td><div align="center"><A HREF="Del_Divisi.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "DivisiID=" & rsDivisi.Fields.Item("DivisiID").Value %>" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; delete record');return document.MM_returnValue">Del</A></div></td>
      <td><div align="center"><A HREF="Mod_Divisi.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "DivisiID=" & rsDivisi.Fields.Item("DivisiID").Value %>"><img src="../Image/modify.gif" width="18" height="18" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; modify record');return document.MM_returnValue"></A></div></td>
      <td><div align="center"><%=(rsDivisi.Fields.Item("DivisiID").Value)%></div></td>
      <td><div align="center"><%=(rsDivisi.Fields.Item("CompanyID").Value)%></div></td>
      <td><div align="center"><%=(rsDivisi.Fields.Item("DivisiName").Value)%></div></td>
      <td><div align="center"><%=(rsDivisi.Fields.Item("UpdateUsr").Value)%></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsDivisi.MoveNext()
Wend
%>
  </table>
  <table border="0" width="600" align="center">
    <tr> 
      <td width="42%" align="center"> <% If MM_offset <> 0 Then %>
        <a href="<%=MM_moveFirst%>"><img src="../Image/first.gif" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; first page');return document.MM_returnValue"></a> 
        <% End If ' end MM_offset <> 0 %> </td>
      <td width="8%" align="center"> <% If MM_offset <> 0 Then %>
        <a href="<%=MM_movePrev%>"><img src="../Image/previous.gif" width="75" height="25" border=0 align="right" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; previous page');return document.MM_returnValue"></a> 
        <% End If ' end MM_offset <> 0 %> </td>
      <td width="8%" align="center"> <% If Not MM_atTotal Then %>
        <a href="<%=MM_moveNext%>"><img src="../Image/next.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt;next page');return document.MM_returnValue"></a> 
        <% End If ' end Not MM_atTotal %> </td>
      <td width="42%" align="center"> <% If Not MM_atTotal Then %>
        <a href="<%=MM_moveLast%>"><img src="../Image/last.gif" width="75" height="25" border=0 align="left" onMouseOver="MM_displayStatusMsg('rz : list divisi -&gt; last page');return document.MM_returnValue"></a> 
        <% End If ' end Not MM_atTotal %> </td>
    </tr>
  </table>
  <p> 
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" name="back" width="105" height="33" align="absmiddle" id="back" onMouseOver="MM_displayStatusMsg('rz : list user -&gt; back to home');MM_displayStatusMsg('rz : list user -&gt; first page');return document.MM_returnValue">
      <param name="BASE" value=".">
      <param name="BGCOLOR" value="">
      <param name="movie" value="back2home.swf">
      <embed src="back2home.swf" width="105" height="33" align="absmiddle" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" name="back" base="." ></embed> 
    </object>
  </p>
  <table width="600" border="1">
    <tr> 
      <td bordercolor="#FFFFCC" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <br>
  <table width="600" border="0">
    <tr>
      <td><div align="center"><a href="MasterBudget.asp" target="_parent">Master Budget</a> | <a href="MasterCompany.asp" target="_parent">Master Company</a> | <a href="MasterCurrency.asp" target="_parent">Master Currency </a> | <a href="MasterDivisi.asp" target="_parent">Master Divisi </a>| <a href="MasterUser.asp" target="_parent">Master User </a> | <a href="MasterVendor.asp" target="_parent">Master Vendor </a></div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsDivisi.Close()
Set rsDivisi = Nothing
%>
