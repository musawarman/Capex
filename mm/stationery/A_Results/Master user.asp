<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/caoexconn.asp" -->
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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "form1") Then

  MM_editConnection = MM_caoexconn_STRING
  MM_editTable = "dbo.UserMS"
  MM_editRedirectUrl = "Master user.asp"
  MM_fieldsStr  = "userid|value|username|value|password|value|select2|value|status|value|select|value"
  MM_columnsStr = "UserID|',none,''|UserName|',none,''|UserPassword|',none,''|UserLEvel|none,none,NULL|UserStatus|',none,''|JabatanID|',none,''"

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
Dim rsuser
Dim rsuser_numRows

Set rsuser = Server.CreateObject("ADODB.Recordset")
rsuser.ActiveConnection = MM_caoexconn_STRING
rsuser.Source = "SELECT JabatanID, UserID, UserName, UserPassword, UserLEvel, UserStatus FROM dbo.UserMS"
rsuser.CursorType = 0
rsuser.CursorLocation = 2
rsuser.LockType = 1
rsuser.Open()

rsuser_numRows = 0
%>
<%
Dim rsjabatan
Dim rsjabatan_numRows

Set rsjabatan = Server.CreateObject("ADODB.Recordset")
rsjabatan.ActiveConnection = MM_caoexconn_STRING
rsjabatan.Source = "SELECT JabatanID FROM dbo.Jabatan"
rsjabatan.CursorType = 0
rsjabatan.CursorLocation = 2
rsjabatan.LockType = 1
rsjabatan.Open()

rsjabatan_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Adduser</title>
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
//-->
</script>
</head>

<body>
<div align="center"> 
  <p> 
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="1000" height="153" align="left">
      <param name="movie" value="../Animasi/baner.swf">
      <param name="quality" value="high"><param name="SCALE" value="exactfit">
      <embed src="../Animasi/baner.swf" width="1000" height="153" align="left" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" scale="exactfit"></embed></object>
  </p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p class="style2">&nbsp;</p>
  <p class="style2">.:: Master User ::. </p>
  <p class="style2">&nbsp; </p>
  <p align="left" class="style2"> 
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="50" height="37">
      <param name="BGCOLOR" value="">
      <param name="movie" value="../Animasi/button23.swf">
      <param name="quality" value="high">
      <embed src="../Animasi/button23.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="50" height="37" ></embed> 
    </object>
    <span class="style4">Back to System Manager </span> </p>
  <table width="1000" border="1">
    <tr>
      <td width="248" height="23" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <p align="center" class="style1">&nbsp;</p>
  <form name="form1" method="POST" action="<%=MM_editAction%>">
    <table width="553" border="1" bordercolor="#3300FF" background="../Image/backg.gif">
      <tr> 
        <td width="154"><div align="center"> 
            <h3><strong>ID</strong></h3>
          </div></td>
        <td width="330"><div align="center"> 
            <h3><strong>User Information</strong></h3>
          </div></td>
      </tr>
      <tr> 
        <td><div align="right"><font color="#FF0000" face="Georgia, Times New Roman, Times, serif">User 
            ID *</font></div></td>
        <td><input name="userid" type="text" id="userid3"></td>
      </tr>
      <tr> 
        <td><div align="right"><font color="#FF0000" face="Georgia, Times New Roman, Times, serif">User 
            Name *</font></div></td>
        <td><input name="username" type="text" id="username"></td>
      </tr>
      <tr> 
        <td><div align="right"><font color="#FF0000" face="Georgia, Times New Roman, Times, serif">User 
            Password *</font></div></td>
        <td><input name="password" type="text" id="password"></td>
      </tr>
      <tr> 
        <td><div align="right"><font color="#FF0000" face="Georgia, Times New Roman, Times, serif">User 
            Level *</font></div></td>
        <td><select name="select2" size="1">
            <option value="1" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("1" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>1</option>
            <option value="2" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("2" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>2</option>
            <option value="3" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("3" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>3</option>
            <option value="4" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("4" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>4</option>
            <option value="5" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("5" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>5</option>
            <option value="6" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("6" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>6</option>
            <option value="7" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("7" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>7</option>
            <option value="8" <%If (Not isNull((rsuser.Fields.Item("UserLEvel").Value))) Then If ("8" = CStr((rsuser.Fields.Item("UserLEvel").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>8</option>
          </select> </td>
      </tr>
      <tr> 
        <td><div align="right"><font color="#FF0000" face="Georgia, Times New Roman, Times, serif">User 
            Status *</font></div></td>
        <td><input name="status" type="text" id="status"> <font color="#009900">A 
          : Active / N : Nonactive</font></td>
      </tr>
      <tr> 
        <td><div align="right"><font color="#FF0000" face="Georgia, Times New Roman, Times, serif">Jabatan 
            ID *</font></div></td>
        <td><select name="select" size="1">
            <%
While (NOT rsjabatan.EOF)
%>
            <option value="<%=(rsjabatan.Fields.Item("JabatanID").Value)%>" <%If (Not isNull((rsjabatan.Fields.Item("JabatanID").Value))) Then If (CStr(rsjabatan.Fields.Item("JabatanID").Value) = CStr((rsjabatan.Fields.Item("JabatanID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsjabatan.Fields.Item("JabatanID").Value)%></option>
            <%
  rsjabatan.MoveNext()
Wend
If (rsjabatan.CursorType > 0) Then
  rsjabatan.MoveFirst
Else
  rsjabatan.Requery
End If
%>
          </select></td>
      </tr>
      <tr> 
        <td><input type="submit" name="Submit" value="Submit"> <input type="reset" name="Submit2" value="Reset"></td>
        <td><div align="right">Tanggal <%=date%></div></td>
      </tr>
      <tr> 
        <td colspan="2"><font color="#0000FF">*) harus diisi</font></td>
      </tr>
    </table>
    <input type="hidden" name="MM_insert" value="form1">
  </form>
  <p>[ <a href="browse%20User.asp">Tampilkan User</a> ]</p>
  <table width="1000" border="1">
    <tr>
      <td bordercolor="#FFFFCC" bgcolor="#669900">&nbsp;</td>
    </tr>
  </table>
  <p align="center" class="style1"><a href="Master%20company.asp">Master Company</a> 
    | <a href="master%20vendor.asp">Master Vendor</a> | <a href="master%20currancy.asp">Master 
    Currency</a> | <a href="master%20budget.asp">Master Budget </a>| <a href="Master%20user.asp">Master 
    User</a></p>
</div>
</body>
</html>
<%
rsuser.Close()
Set rsuser = Nothing
%>
<%
rsjabatan.Close()
Set rsjabatan = Nothing
%>
