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

  MM_editConnection = MM_CapexConn_STRING
  MM_editTable = "dbo.CapexHD"
  MM_editRedirectUrl = "CreateCapexD.asp"
  MM_fieldsStr  = "nocapex|value|divisiid|value|BudgetID|value|PriorityID|value|Tanggal Capex|value|lokasi|value|ammcurrency|value|ammountrate|value|ammidr|value|reason|value|Status|value|Currid|value|hiddenField|value"
  MM_columnsStr = "NoCapex|',none,''|DivisiID|',none,''|BudgetID|',none,''|PriorityID|',none,''|TglCapex|',none,NULL|Lokasi|',none,''|AmmountCurrency|none,none,NULL|AmmountRate|none,none,NULL|AmmountIDR|none,none,NULL|Reason|',none,''|StatusCapex|',none,''|CurrencyID|',none,''|UpdateUsr|',none,''"

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

Dim GetRunNumber__Unit
GetRunNumber__Unit = ""
if(session("Updatedivisi") <> "") then GetRunNumber__Unit = session("Updatedivisi")

%>
<%
Dim rsPriority
Dim rsPriority_numRows

Set rsPriority = Server.CreateObject("ADODB.Recordset")
rsPriority.ActiveConnection = MM_CapexConn_STRING
rsPriority.Source = "SELECT *  FROM dbo.Priority"
rsPriority.CursorType = 0
rsPriority.CursorLocation = 2
rsPriority.LockType = 1
rsPriority.Open()

rsPriority_numRows = 0
%>
<%
Dim rsdivisi
Dim rsdivisi_numRows

Set rsdivisi = Server.CreateObject("ADODB.Recordset")
rsdivisi.ActiveConnection = MM_CapexConn_STRING
rsdivisi.Source = "SELECT * FROM dbo.Divisi"
rsdivisi.CursorType = 0
rsdivisi.CursorLocation = 2
rsdivisi.LockType = 1
rsdivisi.Open()

rsdivisi_numRows = 0
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
<%
Dim rsCurrency
Dim rsCurrency_numRows

Set rsCurrency = Server.CreateObject("ADODB.Recordset")
rsCurrency.ActiveConnection = MM_CapexConn_STRING
rsCurrency.Source = "SELECT * FROM dbo.Currency"
rsCurrency.CursorType = 0
rsCurrency.CursorLocation = 2
rsCurrency.LockType = 1
rsCurrency.Open()

rsCurrency_numRows = 0
%>
<%
Dim rscapexHD__MMColParam
rscapexHD__MMColParam = "1"
If (Request.QueryString("NoCapex") <> "") Then 
  rscapexHD__MMColParam = Request.QueryString("NoCapex")
End If
%>
<%
Dim rscapexHD
Dim rscapexHD_numRows

Set rscapexHD = Server.CreateObject("ADODB.Recordset")
rscapexHD.ActiveConnection = MM_CapexConn_STRING
rscapexHD.Source = "SELECT * FROM dbo.CapexHD WHERE NoCapex = '" + Replace(rscapexHD__MMColParam, "'", "''") + "'"
rscapexHD.CursorType = 0
rscapexHD.CursorLocation = 2
rscapexHD.LockType = 1
rscapexHD.Open()

rscapexHD_numRows = 0
%>
<%

set GetRunNumber = Server.CreateObject("ADODB.Command")
GetRunNumber.ActiveConnection = MM_CapexConn_STRING
GetRunNumber.CommandText = "dbo.P_CreateAutoNumber"
GetRunNumber.Parameters.Append GetRunNumber.CreateParameter("@RETURN_VALUE", 3, 4)
GetRunNumber.Parameters.Append GetRunNumber.CreateParameter("@Unit", 200, 1,5,GetRunNumber__Unit)
GetRunNumber.Parameters.Append GetRunNumber.CreateParameter("@nocapex", 200, 2,20)
GetRunNumber.CommandType = 4
GetRunNumber.CommandTimeout = 0
'GetRunNumber.Prepared = false
GetRunNumber.Execute()

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>:: Create Capex Header ::</title>
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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_validateForm() { //v4.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (isNaN(val)) errors+='- '+nm+' harus berupa angka.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') errors += '- '+nm+' tidak boleh kosong.\n'; }
  } if (errors) alert('Silahkan isikan data dengan lengkap:\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
</head>

<body background="../Image/bgact.gif" onLoad="MM_displayStatusMsg('Create Capex');return document.MM_returnValue;MM_preloadImages('../Image/CreateCapex_on.gif','../Image/ApprovalCapex_on.gif','../Image/CreateAOC_on.gif','../Image/Approvalaoc_on.gif','../Image/holdcapex_on.gif','../Image/holdAoc_on.gif','../Image/Estimation_on.gif','../Image/Actual_on.gif')">
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
            <strong>CREATE CAPEX</strong></font></div></td>
      </tr>
      <tr> 
        <td><div align="left"><img src="../Image/spacer.jpg" width="200" height="2"></div></td>
      </tr>
      <tr> 
        <td height="215"> <div align="center"> 
            <table width="800" border="0">
              <tr> 
                <td height="25"><div align="center"><img src="../../img/garis1.gif" width="600" height="1"></div></td>
              </tr>
            </table>
            <p><img src="../../img/garis1.gif" width="400" height="1"></p>
            <form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1" onSubmit="MM_validateForm('nocapex2','','R','divisiid2','','R','Tanggal Capex2','','R','lokasi2','','R','ammcurrency2','','R','ammountrate2','','RisNum','ammidr2','','RisNum','textarea','','R');return document.MM_returnValue">
              <table width="740" border="1" align="center" bgcolor="#6699FF">
                <tr> 
                  <td height="23"> <div align="center"><font color="#FFFF00">Capex 
                      Header</font></div></td>
                </tr>
                <tr> 
                  <td height="574"> <table width="610" border="1" align="center" bgcolor="#CCCCCC">
                      <tr> 
                        <td width="199"><div align="right"><font color="#0000FF" size="1">No 
                            Capex *</font></div></td>
                        <td width="395"><div align="left"> 
                            <input name="nocapex" id="nocapex2" value="<%= GetRunNumber.Parameters.Item("@nocapex").Value %>" readonly="text"    >
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Divisi 
                            ID *</font></div></td>
                        <td><div align="left"> 
                            <input name="divisiid" readonly="text" id="divisiid2" value="<%= Session("Updatedivisi") %>">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Budget 
                            ID *</font></div></td>
                        <td><div align="left"> 
                            <select name="BudgetID" size="1" id="select">
                              <%
While (NOT rsBudget.EOF)
%>
                              <option value="<%=(rsBudget.Fields.Item("BudgetID").Value)%>" <%If (Not isNull((rsBudget.Fields.Item("BudgetID").Value))) Then If (CStr(rsBudget.Fields.Item("BudgetID").Value) = CStr((rsBudget.Fields.Item("BudgetID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsBudget.Fields.Item("BudgetID").Value)%></option>
                              <%
  rsBudget.MoveNext()
Wend
If (rsBudget.CursorType > 0) Then
  rsBudget.MoveFirst
Else
  rsBudget.Requery
End If
%>
                            </select>
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Priority 
                            ID *</font></div></td>
                        <td><div align="left"> 
                            <select name="PriorityID" size="1" id="select2">
                              <%
While (NOT rsPriority.EOF)
%>
                              <option value="<%=(rsPriority.Fields.Item("PriorityID").Value)%>" <%If (Not isNull((rsPriority.Fields.Item("PriorityID").Value))) Then If (CStr(rsPriority.Fields.Item("PriorityID").Value) = CStr((rsPriority.Fields.Item("PriorityID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsPriority.Fields.Item("PriorityID").Value)%></option>
                              <%
  rsPriority.MoveNext()
Wend
If (rsPriority.CursorType > 0) Then
  rsPriority.MoveFirst
Else
  rsPriority.Requery
End If
%>
                            </select>
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Tanggal 
                            Capex *</font></div></td>
                        <td><div align="left"> 
                            <input name="Tanggal Capex" readonly="text" id="Tanggal Capex2" value="<%=date%>">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Lokasi 
                            * </font></div></td>
                        <td><div align="left"> 
                            <input name="lokasi" type="text" id="lokasi2">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Ammount 
                            Currency *</font></div></td>
                        <td><div align="left"> 
                            <input name="ammcurrency" type="text" id="ammcurrency2">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Ammount 
                            Rate *</font></div></td>
                        <td><div align="left"> 
                            <input name="ammountrate" type="text" id="ammountrate2">
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Ammount 
                            IDR *</font></div></td>
                        <td><div align="left"> 
                            <input name="ammidr" type="text" id="ammidr2">
                          </div></td>
                      </tr>
                      <tr> 
                        <td height="186"><div align="right"><font color="#0000FF" size="1">Reason 
                            * </font></div></td>
                        <td> <div align="left"> 
                            <textarea name="reason" cols="50" rows="10" id="textarea"></textarea>
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Status 
                            Capex *</font></div></td>
                        <td><div align="left"> 
                            <select name="Status" size="1" id="select3">
                              <option value="Unapprove" <%'If (Not isNull((rscapexHD.Fields.Item("StatusCapex").Value))) Then If ("Unapprove" = CStr((rscapexHD.Fields.Item("StatusCapex").Value))) Then Response.Write("SELECTED") : Response.Write("")%> <%If (Not isNull((rscapexHD.Fields.Item("StatusCapex").Value))) Then If ("Unapprove" = CStr((rscapexHD.Fields.Item("StatusCapex").Value))) Then Response.Write("SELECTED") : Response.Write("")%>>Unapprove</option>
                              <%
While (NOT rscapexHD.EOF)
%>
                              <option value="<%=(rscapexHD.Fields.Item("StatusCapex").Value)%>" <%If (Not isNull((rscapexHD.Fields.Item("StatusCapex").Value))) Then If (CStr(rscapexHD.Fields.Item("StatusCapex").Value) = CStr((rscapexHD.Fields.Item("StatusCapex").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rscapexHD.Fields.Item("StatusCapex").Value)%></option>
                              <%
  rscapexHD.MoveNext()
Wend
If (rscapexHD.CursorType > 0) Then
  rscapexHD.MoveFirst
Else
  rscapexHD.Requery
End If
%>
                            </select>
                          </div></td>
                      </tr>
                      <tr> 
                        <td><div align="right"><font color="#0000FF" size="1">Currency 
                            ID *</font></div></td>
                        <td><div align="left"> 
                            <select name="Currid" size="1" id="select4">
                              <%
While (NOT rsCurrency.EOF)
%>
                              <option value="<%=(rsCurrency.Fields.Item("CurrencyID").Value)%>" <%If (Not isNull((rsCurrency.Fields.Item("CurrencyID").Value))) Then If (CStr(rsCurrency.Fields.Item("CurrencyID").Value) = CStr((rsCurrency.Fields.Item("CurrencyID").Value))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(rsCurrency.Fields.Item("CurrencyID").Value)%></option>
                              <%
  rsCurrency.MoveNext()
Wend
If (rsCurrency.CursorType > 0) Then
  rsCurrency.MoveFirst
Else
  rsCurrency.Requery
End If
%>
                            </select>
                          </div></td>
                      </tr>
                      <tr> 
                        <td height="28"> <div align="center"> 
                            <input type="submit" name="Submit" value="Submit">
                            <input type="reset" name="Submit2" value="Reset">
                          </div></td>
                        <td><div align="right"><font color="#663300"> 
                            <input name="hiddenField" type="hidden" value="<%= Session("UpdateUsr") %>">
                            Tanggal <%=date%></font></div></td>
                      </tr>
                      <tr> 
                        <td height="16" colspan="2"> <div align="left">*) harus 
                            diisi</div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <input type="hidden" name="MM_insert" value="form1">
            </form>
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
                    HOME</font></font><font color="#006600">&nbsp; </font>| <a href="Sman.asp"><font color="#006600">SYSTEM 
                    MANAGER</font></a><font color="#006600"> </font>| <a href="Activity.asp"><font color="#006600">ACTIVITIES</font></a><font color="#006600"> 
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
</body>
</html>
<%
rsPriority.Close()
Set rsPriority = Nothing
%>
<%
rsdivisi.Close()
Set rsdivisi = Nothing
%>
<%
rsBudget.Close()
Set rsBudget = Nothing
%>
<%
rsCurrency.Close()
Set rsCurrency = Nothing
%>
<%
rscapexHD.Close()
Set rscapexHD = Nothing
%>













