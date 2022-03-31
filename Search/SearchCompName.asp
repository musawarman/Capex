
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
Dim rsCompany__MMColParam
rsCompany__MMColParam = "1"
If (Request.Form("CompanyName") <> "") Then 
  rsCompany__MMColParam = Request.Form("CompanyName")
End If
%>
<%
Dim rsCompany
Dim rsCompany_numRows

Set rsCompany = Server.CreateObject("ADODB.Recordset")
rsCompany.ActiveConnection = MM_CapexConn_STRING
rsCompany.Source = "SELECT * FROM dbo.Company WHERE CompanyName = '" + Replace(rsCompany__MMColParam, "'", "''") + "' ORDER BY UpdateDate ASC"
rsCompany.CursorType = 0
rsCompany.CursorLocation = 2
rsCompany.LockType = 1
rsCompany.Open()

rsCompany_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
rsCompany_numRows = rsCompany_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsCompany_total
Dim rsCompany_first
Dim rsCompany_last

' set the record count
rsCompany_total = rsCompany.RecordCount

' set the number of rows displayed on this page
If (rsCompany_numRows < 0) Then
  rsCompany_numRows = rsCompany_total
Elseif (rsCompany_numRows = 0) Then
  rsCompany_numRows = 1
End If

' set the first and last displayed record
rsCompany_first = 1
rsCompany_last  = rsCompany_first + rsCompany_numRows - 1

' if we have the correct record count, check the other stats
If (rsCompany_total <> -1) Then
  If (rsCompany_first > rsCompany_total) Then
    rsCompany_first = rsCompany_total
  End If
  If (rsCompany_last > rsCompany_total) Then
    rsCompany_last = rsCompany_total
  End If
  If (rsCompany_numRows > rsCompany_total) Then
    rsCompany_numRows = rsCompany_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsCompany_total = -1) Then

  ' count the total records by iterating through the recordset
  rsCompany_total=0
  While (Not rsCompany.EOF)
    rsCompany_total = rsCompany_total + 1
    rsCompany.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsCompany.CursorType > 0) Then
    rsCompany.MoveFirst
  Else
    rsCompany.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsCompany_numRows < 0 Or rsCompany_numRows > rsCompany_total) Then
    rsCompany_numRows = rsCompany_total
  End If

  ' set the first and last displayed record
  rsCompany_first = 1
  rsCompany_last = rsCompany_first + rsCompany_numRows - 1
  
  If (rsCompany_first > rsCompany_total) Then
    rsCompany_first = rsCompany_total
  End If
  If (rsCompany_last > rsCompany_total) Then
    rsCompany_last = rsCompany_total
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

Set MM_rs    = rsCompany
MM_rsCount   = rsCompany_total
MM_size      = rsCompany_numRows
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
rsCompany_first = MM_offset + 1
rsCompany_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (rsCompany_first > MM_rsCount) Then
    rsCompany_first = MM_rsCount
  End If
  If (rsCompany_last > MM_rsCount) Then
    rsCompany_last = MM_rsCount
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
<title>:: List Company ::</title>
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
.style11 {
	color: #FFFFFF;
	font-weight: bold;
	font-family: verdana;
	font-size: 12px;
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
      <td colspan="2"><div align="center"> 
          <h3><font color="#6699FF">.:: List Company ::. </font></h3>
        </div></td>
    </tr>
    <tr>
      <td colspan="2"><form name="form1" method="post" action="SearchCompName.asp">
          <strong><font color="#FF6600">Look For Company Name : </font></strong> 
          <input name="CompanyName" type="text" id="CompanyName">
          <input name="Search" type="submit" id="Search" value="Search">
        </form> </td>
    </tr>
    <tr bgcolor="#669900"> 
      <td height="14" colspan="2"> <div align="center"><span class="style7">Welcome 
          <%= Session("UpdateUsr") %></span></div></td>
    </tr>
    <tr bgcolor="#FF9900"> 
      <td height="14" colspan="2"> <div align="center"><font color="#0000A0"><strong>..:: 
          Records <%=(rsCompany_first)%> to <%=(rsCompany_last)%> of <%=(rsCompany_total)%> ::..</strong></font> </div></td>
    </tr>
  </table>
  <br>
  <% If Not rsCompany.EOF Or Not rsCompany.BOF Then %>
  <table border="1" align="center" cellspacing="0" bordercolor="#FFFFFF">
    <tr bgcolor="#6666FF"> 
      <td width="22"><div align="center"><span class="style5">No</span></div></td>
      <td width="36"><div align="center"><span class="style5">Delete</span></div></td>
      <td width="40"><div align="center"><span class="style5">Update</span></div></td>
      <td width="149"><div align="center"><span class="style5">CompanyID</span></div></td>
      <td width="189"><div align="center"><span class="style5">CompanyName</span></div></td>
      <td width="138"><div align="center"><span class="style5">Post By </span></div></td>
    </tr>
    <% While ((Repeat1__numRows <> 0) AND (NOT rsCompany.EOF)) %>
    <tr bgcolor="#CCCCCC"> 
      <td height="26"><div align="center"><%=(Repeat1__index + 1)%>&nbsp;</div></td>
      <td><div align="center"><A HREF="../MainMenu/Del_Company.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "CompanyID=" & rsCompany.Fields.Item("CompanyID").Value %>" onMouseOver="MM_displayStatusMsg('rz : list company -&gt; delete record');return document.MM_returnValue">Del</A></div></td>
      <td><div align="center"><A HREF="../MainMenu/Mod_Company.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "CompanyID=" & rsCompany.Fields.Item("CompanyID").Value %>"><img src="../Image/modify.gif" width="18" height="18" onMouseOver="MM_displayStatusMsg('rz : list company -&gt; modify record');return document.MM_returnValue"></A></div></td>
      <td><div align="center"><%=(rsCompany.Fields.Item("CompanyID").Value)%></div></td>
      <td><div align="center"><%=(rsCompany.Fields.Item("CompanyName").Value)%></div></td>
      <td><div align="center"><%=(rsCompany.Fields.Item("UpdateUsr").Value)%></div></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsCompany.MoveNext()
Wend
%>
  </table>
  <% End If ' end Not rsCompany.EOF Or NOT rsCompany.BOF %>
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
rsCompany.Close()
Set rsCompany = Nothing
%>