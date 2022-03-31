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
' *** Redirect if username exists
MM_flag="MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  MM_dupKeyRedirect="MasterBudgetCek.asp"
  MM_rsKeyConnection=MM_CapexConn_STRING
  MM_dupKeyUsernameValue = CStr(Request.Form("budgetID"))
  MM_dupKeySQL="SELECT BudgetID FROM dbo.CategoryBudget WHERE BudgetID='" & MM_dupKeyUsernameValue & "'"
  MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then 
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_CapexConn_STRING
  MM_editTable = "dbo.CategoryBudget"
  MM_editRedirectUrl = "MasterBudget.asp"
  MM_fieldsStr  = "budgetID|value|budgetname|value|totalbudget|value|tanggalbudget|value|SaldoBudget|value|hiddenField|value"
  MM_columnsStr = "BudgetID|',none,''|BudgetName|',none,''|TotalBudget|none,none,NULL|TanggalBudget|',none,NULL|SaldoBudget|none,none,NULL|UpdateUsr|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
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
Dim rsBudget
Dim rsBudget_numRows

Set rsBudget = Server.CreateObject("ADODB.Recordset")
rsBudget.ActiveConnection = MM_CapexConn_STRING
rsBudget.Source = "SELECT * FROM dbo.CategoryBudget"
rsBudget.CursorType = 0
rsBudget.CursorLocation = 2
rsBudget.LockType = 1
rsBudget.Open()

rsBudget_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Master Budget ::</title>
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

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style5 {font-weight: bold}
.style8 {font-family: verdana}
.style9 {color: #FF0000}
.style10 {font-family: verdana; color: #FF0000; }
.style12 {color: #0000FF}
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
      <td colspan="2"><h3 align="center"><font color="#6699FF">.:: Master Budget 
          ::. </font></h3></td>
    </tr>
    <tr> 
      <td width="457" height="39"> <div align="left"> 
          <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="50" height="37" hspace="10" align="absmiddle" onMouseOver="MM_displayStatusMsg('rz : delete vendor confirmation -&gt; back to system manager');return document.MM_returnValue">
            <param name="BASE" value=".">
            <param name="BGCOLOR" value="">
            <param name="movie" value="back2sm.swf">
            <embed src="back2sm.swf" width="50" height="37" hspace="10" align="absmiddle" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" base="." ></embed> 
          </object>
          <span class="style4"><font color="#0000FF" size="-1" face="Arial, Helvetica, sans-serif">back 
          to system manager</font> </span> </div></td>
      <td width="120"> 
        <div align="center"><a href="../Search/Search.asp"></a><font size="2"><a href="<%= MM_Logout %>" target="_parent">Logout</a> 
          |<strong> <a href="../Search/Search.asp" target="_parent">Search</a></strong></font></div></td>
    </tr>
  </table>
  <table width="600" border="1">
    <tr>
      <td width="248" height="16" bgcolor="#669900"> <div align="center"><span class="style5"><font color="#FFFFFF">W</font><font color="#FFFFFF">elcome 
          <%= Session("UpdateUsr") %> </font></span></div></td>
    </tr>
  </table>
  <form name="form1" method="POST" action="<%=MM_editAction%>">
    <table width="525" border="1" cellspacing="0" bordercolor="#999999">
      <tr bgcolor="#000099"> 
        <td colspan="2"> 
          <div align="center"> <font color="#FFFFFF" size="3" face="Comic Sans MS">Budget 
            Category Information</font></div>          </td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td width="161"> 
          <div align="right" class="style8 style9"><font color="#FF0000">Budget 
            ID </font>: </div></td>
        <td width="348"> 
          <div align="left"> 
            <input name="budgetID" type="text" id="budgetID">
          *</div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right" class="style10">Budget Name : </div></td>
        <td> 
          <div align="left"> 
            <input name="budgetname" type="text" id="budgetname" size="50">
            *
          </div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right" class="style10">Total Budget : </div></td>
        <td> 
          <div align="left"> 
            <input name="totalbudget" type="text" id="totalbudget">
            *
          </div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td height="26"> 
          <div align="right" class="style10">Tanggal Budget : </div></td>
        <td> 
          <div align="left"> 
            <input name="tanggalbudget" type="text" id="tanggalbudget"value="<%=date%>" >
            *
            <font color="#0000FF">format (mm/dd/yyyy)</font></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="right" class="style10">Saldo Budget : </div></td>
        <td> 
          <div align="left"> 
            <input name="SaldoBudget" type="text" id="SaldoBudget">
            *
          </div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td> 
          <div align="center"> 
            <input name="Submit" type="submit" onClick="MM_validateForm('budgetID','','R','budgetname','','R','totalbudget','','RisNum','tanggalbudget','','R','SaldoBudget','','RisNum');return document.MM_returnValue" onMouseOver="MM_displayStatusMsg('rz : master budget -&gt; submit&lt;add new record&gt;');return document.MM_returnValue" value="Submit">
            <input name="Submit2" type="reset" onMouseOver="MM_displayStatusMsg('rz : master budget -&gt; reset');return document.MM_returnValue" value="Reset">
          </div></td>
        <td> 
          <div align="right"> 
            <input name="hiddenField" type="hidden" value="<%= Session("UpdateUsr") %>">
            <span class="style12">Date:</span> <%=date%></div></td>
      </tr>
      <tr bgcolor="#CCCCCC"> 
        <td colspan="2"> 
          <div align="left"><font color="#0000FF">*) required </font></div></td>
      </tr>
    </table>
    [ <a href="ListBudget.asp" target="_parent">List Budget </a>]<br>
    <input type="hidden" name="MM_insert" value="form1">
  </form>
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" name="back" width="105" height="33" align="absmiddle" id="back" onMouseOver="MM_displayStatusMsg('rz : list user -&gt; back to home');MM_displayStatusMsg('rz : list user -&gt; first page');return document.MM_returnValue">
    <param name="BASE" value=".">
    <param name="BGCOLOR" value="">
    <param name="movie" value="back2home.swf">
    <embed src="back2home.swf" width="105" height="33" align="absmiddle" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" name="back" base="." ></embed> 
  </object>
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
rsBudget.Close()
Set rsBudget = Nothing
%>
