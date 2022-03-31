
<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/CapexConn.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers="8"
MM_authFailedURL="failed.asp"
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
Dim rsCurrency__MMColParam
rsCurrency__MMColParam = "1"
If (Request.Form("CurrencyName") <> "") Then 
  rsCurrency__MMColParam = Request.Form("CurrencyName")
End If
%>
<%
Dim rsCurrency
Dim rsCurrency_numRows

Set rsCurrency = Server.CreateObject("ADODB.Recordset")
rsCurrency.ActiveConnection = MM_CapexConn_STRING
rsCurrency.Source = "SELECT * FROM dbo.Currency WHERE CurrencyName = '" + Replace(rsCurrency__MMColParam, "'", "''") + "' ORDER BY UpdateDate ASC"
rsCurrency.CursorType = 0
rsCurrency.CursorLocation = 2
rsCurrency.LockType = 1
rsCurrency.Open()

rsCurrency_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsCurrency_numRows = rsCurrency_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsCurrency_total
Dim rsCurrency_first
Dim rsCurrency_last

' set the record count
rsCurrency_total = rsCurrency.RecordCount

' set the number of rows displayed on this page
If (rsCurrency_numRows < 0) Then
  rsCurrency_numRows = rsCurrency_total
Elseif (rsCurrency_numRows = 0) Then
  rsCurrency_numRows = 1
End If

' set the first and last displayed record
rsCurrency_first = 1
rsCurrency_last  = rsCurrency_first + rsCurrency_numRows - 1

' if we have the correct record count, check the other stats
If (rsCurrency_total <> -1) Then
  If (rsCurrency_first > rsCurrency_total) Then
    rsCurrency_first = rsCurrency_total
  End If
  If (rsCurrency_last > rsCurrency_total) Then
    rsCurrency_last = rsCurrency_total
  End If
  If (rsCurrency_numRows > rsCurrency_total) Then
    rsCurrency_numRows = rsCurrency_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsCurrency_total = -1) Then

  ' count the total records by iterating through the recordset
  rsCurrency_total=0
  While (Not rsCurrency.EOF)
    rsCurrency_total = rsCurrency_total + 1
    rsCurrency.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsCurrency.CursorType > 0) Then
    rsCurrency.MoveFirst
  Else
    rsCurrency.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsCurrency_numRows < 0 Or rsCurrency_numRows > rsCurrency_total) Then
    rsCurrency_numRows = rsCurrency_total
  End If

  ' set the first and last displayed record
  rsCurrency_first = 1
  rsCurrency_last = rsCurrency_first + rsCurrency_numRows - 1
  
  If (rsCurrency_first > rsCurrency_total) Then
    rsCurrency_first = rsCurrency_total
  End If
  If (rsCurrency_last > rsCurrency_total) Then
    rsCurrency_last = rsCurrency_total
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

Set MM_rs    = rsCurrency
MM_rsCount   = rsCurrency_total
MM_size      = rsCurrency_numRows
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
rsCurrency_first = MM_offset + 1
rsCurrency_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsCurrency_first > MM_rsCount) Then
    rsCurrency_first = MM_rsCount
  End If
  If (rsCurrency_last > MM_rsCount) Then
    rsCurrency_last = MM_rsCount
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
<title>:: List Currency ::</title>
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
<link href="../css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style5 {color: #FFFFFF}
.style7 {color: #FFFFFF; font-weight: bold; }
.style8 {
	font-family: verdana;
	font-size: 12px;
	font-weight: bold;
}
-->
</style>
</head>

<body bgcolor="#FFFFFF" background="../Image/bg.gif" onLoad="MM_preloadImages('../Image/over_first%20page.gif','../Image/over_previous%20page.gif','../Image/over_next%20page.gif','../Image/over_last%20page.gif')">
<div align="center"> 
  <table width="756" border="0">
    <tr> 
      <td width="750" colspan="2"><img src="../Image/banner2.gif" width="750" height="100"></td>
    </tr>
    <tr> 
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: List Currency 
          ::. </font></h3></td>
    </tr>
    <tr>
      <td colspan="2"><form name="form1" method="post" action="SearchCurrName.asp">
          <font color="#FF6600"><strong>Look For Currency Name : </strong></font> 
          <input name="CurrencyName" type="text" id="CurrencyName">
          <font color="#FF6600"></font> 
          <input name="Search" type="submit" id="Search" value="Search">
        </form></td>
    </tr>
    <tr bgcolor="#669900"> 
      <td colspan="2"> <div align="center"><span class="style7">Welcome <%= Session("UpdateUsr") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td colspan="2"> <div align="center"><font color="#0000A0"><strong>..:: 
          Records <%=(rsCurrency_first)%> to <%=(rsCurrency_last)%> of <%=(rsCurrency_total)%> ::..</strong></font></div></td>
    </tr>
  </table>
  <br>
  <% If Not rsCurrency.EOF Or Not rsCurrency.BOF Then %>
  <table border="1" align="center" cellspacing="0" bordercolor="#FFFFFF">
    <tr bgcolor="#6666FF"> 
      <td width="19"><div align="center"><span class="style5">No</span></div></td>
      <td width="36"><div align="center"><span class="style5">Delete</span></div></td>
      <td width="40"><div align="center"><span class="style5">Update</span></div></td>
      <td width="158"><div align="center"><span class="style5">CurrencyID</span></div></td>
      <td width="171"><div align="center"><span class="style5">CurrencyName</span></div></td>
      <td width="150"><div align="center" class="style5">Post By </div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsCurrency.EOF)) %>
    <tr bgcolor="#CCCCCC"> 
      <td height="26"><div align="center"><%=(Repeat1__index + 1)%></div></td>
      <td><div align="center"><A HREF="../MainMenu/Del_Currency.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "CurrencyID=" & rsCurrency.Fields.Item("CurrencyID").Value %>" onMouseOver="MM_displayStatusMsg('rz : list currency -&gt; delete record');return document.MM_returnValue">Del</A></div></td>
      <td><div align="center"><A HREF="../MainMenu/Mod_Currency.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "CurrencyID=" & rsCurrency.Fields.Item("CurrencyID").Value %>"><img src="../Image/modify.gif" width="18" height="18" onMouseOver="MM_displayStatusMsg('rz : list currency -&gt; modify record');return document.MM_returnValue"></A></div></td>
      <td><div align="center"><%=(rsCurrency.Fields.Item("CurrencyID").Value)%></div></td>
      <td><div align="center"><%=(rsCurrency.Fields.Item("CurrencyName").Value)%></div></td>
      <td><div align="center"><%=(rsCurrency.Fields.Item("UpdateUsr").Value)%></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsCurrency.MoveNext()
Wend
%>
  </table>
  <% End If ' end Not rsCurrency.EOF Or NOT rsCurrency.BOF %>
  <p>&nbsp; </p>
  <table width="750" border="1">
    <tr> 
      <td bordercolor="#FFFFCC" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <br>
  <table width="600" border="0">
    <tr>
      <td><div align="center"><a href="../MainMenu/MasterBudget.asp" target="_parent">Master Budget</a> | <a href="../MainMenu/MasterCompany.asp" target="_parent">Master Company</a> | <a href="../MainMenu/MasterCurrency.asp" target="_parent">Master Currency </a> | <a href="../MainMenu/MasterDivisi.asp" target="_parent">Master Divisi </a>| <a href="../MainMenu/MasterUser.asp" target="_parent">Master User </a> | <a href="../MainMenu/MasterVendor.asp" target="_parent">Master Vendor </a></div></td>
    </tr>
  </table>
</div>
</body>
</html>
<%
rsCurrency.Close()
Set rsCurrency = Nothing
%>
