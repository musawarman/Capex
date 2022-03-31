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
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>

<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_CapexConn_STRING
  MM_editTable = "dbo.Currency"
  MM_editColumn = "CurrencyID"
  MM_recordId = "'" + Request.Form("MM_recordId") + "'"
  MM_editRedirectUrl = "ListCurrency.asp"
  MM_fieldsStr  = "CurrencyName|value|hiddenField|value"
  MM_columnsStr = "CurrencyName|',none,''|UpdateUsr|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim rsCurrency__MMColParam
rsCurrency__MMColParam = "1"
If (Request.QueryString("CurrencyID") <> "") Then 
  rsCurrency__MMColParam = Request.QueryString("CurrencyID")
End If
%>
<%
Dim rsCurrency
Dim rsCurrency_numRows

Set rsCurrency = Server.CreateObject("ADODB.Recordset")
rsCurrency.ActiveConnection = MM_CapexConn_STRING
rsCurrency.Source = "SELECT * FROM dbo.Currency WHERE CurrencyID = '" + Replace(rsCurrency__MMColParam, "'", "''") + "'"
rsCurrency.CursorType = 0
rsCurrency.CursorLocation = 2
rsCurrency.LockType = 1
rsCurrency.Open()

rsCurrency_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Currency Modification ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-size: 18px}
.style2 {
	font-size: 36px;
	color: #0000FF;
}
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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' is required.\n'; }
  } if (errors) alert('The following error(s) occurred:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style5 {font-weight: bold}
.style6 {font-family: verdana}
.style7 {color: #FF0000}
-->
</style>
<% session.timeout=15%>
</head>

<body background="../Image/bg.gif">
<div align="center"> 
  <table width="600" border="0">
    <tr> 
      <td colspan="2"><img src="../Image/banner2.gif" width="750" height="100"></td>
    </tr>
    <tr> 
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: Currency Modification 
          ::. </font></h3></td>
    </tr>
    <tr> 
      <td width="470"><div align="left"> 
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="50" height="37" hspace="10" align="absmiddle" onMouseOver="MM_displayStatusMsg('rz : delete vendor confirmation -&gt; back to system manager');return document.MM_returnValue">
            <param name="BASE" value=".">
            <param name="BGCOLOR" value="">
            <param name="movie" value="back2sm.swf">
            <embed src="back2sm.swf" width="50" height="37" hspace="10" align="absmiddle" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" base="." ></embed> 
          </object>
          <span class="style4"><font color="#0000FF" size="-1" face="Arial, Helvetica, sans-serif">back 
          to system manager </font></span></div></td>
      <td width="120">
<div align="center"><font size="2"><a href="<%= MM_Logout %>" target="_parent">Logout</a> 
          |<strong> <a href="../Search/Search.asp" target="_parent">Search</a></strong></font></div></td>
    </tr>
  </table>
  <table width="600" border="1">
    <tr>
      <td width="248" height="23" bgcolor="#669900"><div align="center"><span class="style5"><font color="#FFFFFF">W</font><font color="#FFFFFF">elcome 
          <%= Session("UpdateUsr") %> </font></span></div></td>
    </tr>
  </table>
  <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
    <table width="553" border="1" cellspacing="0" bordercolor="#999999">
      <tr bgcolor="#000099"> 
        <td colspan="2"> <div align="center"> <font color="#FFFFFF" size="3" face="Comic Sans MS">Currency 
            Information</font></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td height="25"> 
          <div align="right"><span class="style6"><span class="style7">Currency 
            ID : </span></span></div></td>
        <td width="330"> 
          <div align="left"> <%=(rsCurrency.Fields.Item("CurrencyID").Value)%> </div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right"><span class="style6"><span class="style7">Currency 
            Name : </span></span></div></td>
        <td> 
          <div align="left"> 
            <input value="<%=((rsCurrency.Fields.Item("CurrencyName").Value))%>" name="CurrencyName" type="text" id="CurrencyName">
            *</div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="center"> 
            <input name="Submit" type="submit" onClick="MM_validateForm('CurrencyName','','R');return document.MM_returnValue" onMouseOver="MM_displayStatusMsg('rz : currency modification -&gt; update record');return document.MM_returnValue" value="Update">
            <input name="Submit2" type="reset" onMouseOver="MM_displayStatusMsg('rz : currency modification -&gt; reset');return document.MM_returnValue" value="Reset">
          </div></td>
        <td> 
          <div align="right"> 
            <input name="hiddenField" type="hidden">
            Date: <%=date%></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td colspan="2"> 
          <div align="left">*) required</div></td>
      </tr>
    </table>
    <input type="hidden" name="MM_update" value="form1">
    <input type="hidden" name="MM_recordId" value="<%= rsCurrency.Fields.Item("CurrencyID").Value %>">
  </form>
  <form name="form2" method="post" action="">
    <input name="Submit22" type=button class=btn onclick=history.back() onMouseOver="MM_displayStatusMsg('rz : delete company confirmation -&gt; back to last form');return document.MM_returnValue" value="Cancel">
  </form>
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
rsCurrency.Close()
Set rsCurrency = Nothing
%>
