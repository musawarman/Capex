<% Option Explicit %>
<!--#INCLUDE FILE="defaults.inc" -->
<!--#INCLUDE FILE="LocalizedStrings.asp" -->
<%
 
' Copyright © 2002 Crystal Decisions, Inc.
Response.ExpiresAbsolute = Now() - 1
Session.CodePage = 65001   ' Set to Unicode
' Write the cookie values
Dim m_MainView, m_Viewer

getCookieValues()
setCookieValues()

Sub getCookieValues()
	' First default to the cookie value	
	m_MainView = Request.Cookies("reportListing")("mainView")
	if (m_MainView = "") then m_MainView = mainViewDefault
	m_Viewer = Request.Cookies("reportListing")("viewer")
	if (m_Viewer = "") then m_Viewer = ViewerDefault
	
	if Request.QueryString("mainView") <> "" then _
		m_MainView = Request.QueryString("mainView")
	if Request.QueryString("viewer") <> "" then _
		m_Viewer = Request.QueryString("viewer")
End Sub

Sub setCookieValues()
	Response.Cookies("reportListing").Expires = DateAdd("yyyy", 1, Now)	' Expire one year from now
	Response.Cookies("reportListing")("mainView") = m_MainView
	Response.Cookies("reportListing")("viewer") = m_Viewer
End Sub

Sub UpdatePage()
	if Request.QueryString("viewer") <> "" then _
	    Response.Write "<SCRIPT> window.opener.location.reload(); window.close();  </SCRIPT>" + vbCrLf
End Sub

Sub outputVal(theVar, value)
	if theVar = value then Response.Write " CHECKED "
	Response.Write "value=" + chr(34) + value + chr(34)
End Sub
%>

<html>
<head>
<title><%= L_PREFERENCESTITLE %></title>
<script language="javascript" src="include/cookiescripts.js"></script>

<script language="javascript" src="include/pophelpwindow.js"></script>

<script language="javascript">
var link_path = 'include/';	// not used
function Cancel() { 
	window.close(); 
}
function applyValues() {
	document.forms["Preferences"].submit();
}
</script>
<link href="include/default.css" type="text/css" rel="stylesheet" name="stylelink">
<script language="javascript" src="include/buttons.js"></script>

</head>
<body leftMargin="0" topMargin="0" marginheight="0" marginwidth="0"><!-- HEADER -->
<table class="header" cellSpacing="0" cellPadding="0" width="100%" border="0">
  <tbody>
  <tr>
    <td class="header"><img alt="<%= L_CORPORATELOGO %>" src="include/eportfolio_default.gif" border="0" WIDTH="210" HEIGHT="44"> </td>
    <td class="header" align="right"></td>
    <td class="header" noWrap align="right" colSpan="2">
      <table cellSpacing="0" cellPadding="0" border="0">
        <tbody>
        <tr>
          <td align="middle"></td>
          <td align="middle">&nbsp;</td>
          <td align="middle"><a class="header" href="javascript:applyValues();">
          <img alt="<%= L_OK %>" src="include/apply_up_default.gif" border="0" name="Apply" valign="bottom" WIDTH="32" HEIGHT="32"></a></td>
          <td align="middle">&nbsp;</td>
          <td align="middle"><a class="header" href="javascript:Cancel();"><img alt="<%= L_CLOSE %>" src="include/back_up_default.gif" border="0" name="Close" valign="bottom" WIDTH="32" HEIGHT="32"></a></td>
          <td align="middle">&nbsp;</td>
          <td align="middle"><a class="header" href="javascript:popHelpWindow();"><img alt="<%= L_HELP %>" src="include/help_up_default.gif" border="0" name="Help" valign="bottom" WIDTH="32" HEIGHT="32"></a></td>
          <td align="middle" rowSpan="2">&nbsp;</td></tr>
        <tr>
          <td align="middle"></td>
          <td align="middle"><a class="menuItem">&nbsp;|&nbsp;</a></td>
          <td align="middle"><a class="menuItem" href="javascript:applyValues();"><%= L_OK %></a></td>
          <td align="middle"><a class="menuItem">&nbsp;|&nbsp;</a></td>
          <td align="middle"><a class="menuItem" href="javascript:Cancel();"><%= L_CLOSE %></a></td>
          <td align="middle"><a class="menuItem">&nbsp;|&nbsp;</a></td>
          <td align="middle"><a class="menuItem" href="javascript:popHelpWindow();"><%= L_HELP %></a></td></tr></tbody></table></td></tr>
  <tr>
    <td class="menu" width="64%" colSpan="2">&nbsp;</td>
    <td class="menu" vAlign="top" align="right" width="1%"><img alt="<%= L_CORPORATELOGO %>" src="include/menu_dot_bg_default.gif" WIDTH="19" HEIGHT="56"></td>
    <td class="header" width="35%">&nbsp;</td></tr></tbody></table>
<table class="list" style="BACKGROUND-COLOR: white" cellSpacing="0" cellPadding="3" width="100%" border="0">
  <tbody>
  <tr>
    <td class="list"><span class="listSelected"><%= L_USERPREFERENCES %></span><br>
      <hr SIZE="0">
    </td></tr></tbody></table><br>
<table class="main" cellSpacing="0" cellPadding="3" width="100%" border="0">
  <form name="Preferences" action="preferences.asp" method="get">
  <tbody>
  <tr>
    <td class="main" width="20%"><%= L_MAINDISPLAY %></td>
    <td class="main"><input type="radio" <% outputVal m_MainView, "0" %> name="mainView" id="mainView"><%= L_DISPLAYSEECE %>
     </td></tr>
  <tr>
    <td width="20%">&nbsp;</td>
    <td class="main"><input type="radio" <% outputVal m_MainView, "1" %> name="mainView" id="mainView"><% = L_DISPLAYMINIMAL %></td></tr>
  <tr>
    <td>&nbsp</td>
    <td><font size="1">* <%= L_EPORTFOLIOLITECOMMENT %><br></td>
  <tr>
    <td colSpan="2">
      <hr SIZE="0">
    </td></tr>
  <tr>
    <td class="main" width="20%"><%= L_VIEWTEXT %></td>
<!--    <td class="main"><input type="radio" <% outputVal m_Viewer, "0" %> name="viewer" id="viewer"><%= L_VIEWACTIVEX %></td></tr>
  <tr>
    <td width="20%">&nbsp;</td> -->
    <td class="main"><input type="radio" <% outputVal m_Viewer, "1" %> name="viewer" id="viewer"><%= L_VIEWHTMLPAGE %></td></tr>
  <tr>
    <td width="20%">&nbsp;</td>
    <td class="main"><input type="radio" <% outputVal m_Viewer, "2" %> name="viewer" id="viewer"><%= L_VIEWHTMLINTERACTIVE %></td></tr>
<!--  <tr>
    <td width="20%">&nbsp;</td>
    <td class="main"><input type="radio" <% outputVal m_Viewer, "3" %> name="viewer" id="viewer"><%= L_VIEWJAVA %></td></tr>
-->
  <tr>
    <td width="20%">&nbsp;</td>
    <td class="main"><input type="radio" <% outputVal m_Viewer, "4" %> name="viewer" id="viewer"><%= L_VIEWREPORTPARTS %></td></tr>
  <tr>
    <td class="list" align="right" colSpan="2"><a class="list" href="javascript:applyValues();"><%= L_OK %></a> | <a class="list" href="javascript:Cancel();"><%= L_CLOSE %></a> 
</td></tr></tbody></table></form>

<% updatePage() %>
</body></html>
