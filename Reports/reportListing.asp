<% Option Explicit 
' Copyright © 2002 Crystal Decisions, Inc.
Response.ExpiresAbsolute = Now() - 1
Session.CodePage = 65001   ' Set to Unicode
%>
<!--#INCLUDE FILE="helperFunctions.asp" -->
<meta http-equiv="content-type" content="text/html; charset=utf-8">
<script language="JavaScript1.2" type="text/javascript">
var displayedIObjs;
displayedIObjs = new Array();
var link_path;
link_path = 'include/';
HM_PG_MenuWidth = 120;
HM_PG_FontFamily = "verdana";
HM_PG_FontSize = 8;
HM_PG_FontBold = 0;
HM_PG_FontItalic = 0;
HM_PG_FontColor = "white";
HM_PG_FontColorOver = "black";
HM_PG_BGColor = "#006699"; // darker
HM_PG_BGColorOver = "#7B9AC5"; // lighter
HM_PG_ItemPadding = 1;
HM_PG_BorderWidth = 1;
HM_PG_BorderColor = "black";
HM_PG_BorderStyle = "solid";
HM_PG_SeparatorSize = 0;
HM_PG_SeparatorColor = "red";
HM_PG_ImageSrc = link_path + "images/tri.gif";
HM_PG_ImageSrcLeft = link_path + "images/triL.gif";
HM_PG_ImageSize = 10;
HM_PG_ImageHorizSpace = 0;
HM_PG_ImageVertSpace = 2;
HM_PG_KeepHilite = 0; 
HM_PG_ClickStart = false;
HM_PG_ClickKill = true;
HM_PG_ChildOverlap = 0;
HM_PG_ChildOffset = 10;
HM_PG_ChildPerCentOver = null;
HM_PG_TopSecondsVisible = .5;
HM_PG_StatusDisplayBuild = 0;
HM_PG_StatusDisplayLink = 0;
HM_PG_UponDisplay = null;
HM_PG_UponHide = null;
HM_PG_RightToLeft = false;
HM_PG_CreateTopOnly = 1;
HM_PG_ShowLinkCursor = 1;
HM_PG_NSFontOver = true;
//HM_a_TreesToBuild = [1,2];
//-->
</script>
<link href="include/default.css" type="text/css" rel="stylesheet" name="stylelink">
<script language="javascript" src="include/cookiescripts.js"></script>

<script language="javascript" src="include/pophelpwindow.js"></script>

<script language="javascript" src="include/KeyDownEvent.js"></script>

<script language="javascript">
// SUBMIT THE SEARCH CRITERIA IF USER HAS NOT ENTERED BLANK STRING
function search() {
}


// SEND FOCUS TO THE SEARCH STRING TEXT BOX
function moveFocus() {
	document.searchForm.searchString.focus();
	document.searchForm.searchString.select();
}
// RELOAD PAGE WITH NEW OBJECT TYPE
function selectObjectType( objTypeElement ) {
	document.obtypeform.otype.value = objTypeElement.options[objTypeElement.selectedIndex].value;
	document.obtypeform.submit();
}
// LAUNCH FEATURE NOT IN VERSION DIALOG WINDOW
function NotInVersion() {
	document.location = link_path + "not_in_std.htm";
}
// DUMMY FUNCTION DOES NOTHING.  USED BY LINKS WITH NO REAL HREF.
function doNothing() {}

function launchPreferences()
{
	var features = 'scrollbars=yes,status=no,location=no,resizable=yes,width=600,height=435,left=150,top=100';
	window.open("Preferences.asp","",features);
}

// DISABLE RIGHT CLICK IN NETSCAPE 4
function norightclick(e) 
{
	if (window.Event) 
	{
		if (e.which == 2 || e.which == 3)
			return false;
		else
			routeEvent(e);
	}
}
if (document.layers)
{
	window.captureEvents(Event.MOUSEDOWN);
	window.onmousedown = norightclick;
}
</script>

<script language="javascript" src="include/buttons.js"></script>

<script></script>

<style type="text/css">UL.list {
	MARGIN-LEFT: 18px; LIST-STYLE-TYPE: square
}
</style>

</head>
<body leftMargin="0" topMargin="0" marginwidth="0" marginheight="0">
<div id="overDiv" style="Z-INDEX: 1; POSITION: absolute"></div>
<script language="JavaScript" src="include/overlib.js"></script>

<div align="left"><!-- HEADER -->
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
  <tbody>
  <tr>
    <td class="header" vAlign="center" colSpan="3">
      <table cellSpacing="0" cellPadding="0" width="100%" border="0">
        <tbody>
        <tr>
          <td class="header" vAlign="center"><img alt="<%= L_CORPORATELOGO %>" src="include/eportfolio_default.gif" border="0" WIDTH="210" HEIGHT="44"> </td>
          <td class="header" align="right"></td>
          <td class="header" noWrap align="right">
            <table cellSpacing="0" cellPadding="0" border="0">
              <tbody>
              <tr>
                <td align="middle"><% if m_MainView = "0" then %><img alt="<%= L_UPGRADEALERTS %>" src="include/alerts_up_default.gif" border="0" name="Alerts" valign="bottom" WIDTH="32" HEIGHT="32"><% end if %></td>
                <td align="middle">&nbsp;</td>
                <td align="middle"><% if m_MainView = "0" then %><img alt="<%= L_UPGRADEFAVORITES %> " src="include/favorites_up_default.gif" border="0" name="Favorites" valign="bottom" WIDTH="32" HEIGHT="32"><% end if %></td>
                <td align="middle">&nbsp;</td>
                <td align="middle"><% if m_MainView = "0" then %><img alt="<%= L_UPGRADEFOLDERS %>" src="include/organize_up_default.gif" border="0" name="Organize" valign="bottom" WIDTH="32" HEIGHT="32"><% end if %></td>
                <td align="middle">&nbsp;</td>
                <td align="middle"><a href="javascript:launchPreferences();"><img alt="<%= L_PREFERENCES %>" src="include/preferences_up_default.gif" border="0" name="Settings" valign="bottom" WIDTH="32" HEIGHT="32"></a></td>
                <td align="middle">&nbsp;</td>
                <td align="middle"><% if m_MainView = "0" then %><img alt="<%= L_UPGRADESECURITY %>" src="include/logoff_up_default.gif" border="0" name="Logoff" valign="bottom" WIDTH="32" HEIGHT="32"><% end if %></td>
                <td align="middle">&nbsp;</td>
                <td align="middle"><a href="javascript:popHelpWindow();"><img alt="<%= L_HELP %>" src="include/help_up_default.gif" border="0" name="Help" valign="bottom" WIDTH="32" HEIGHT="32"></a></td>
                <td align="middle" rowSpan="2">&nbsp;</td></tr>
              <tr>
                <td align="middle"><% if m_MainView = "0" then %><span class="listUnavailable"><%= L_ALERTS %></a></span><% end if %></td>
                <td align="middle"><% if m_MainView = "0" then %><a class="menuItem">&nbsp;|&nbsp;</a><% end if %></td>
                <td align="middle"><% if m_MainView = "0" then %><span class="listUnavailable"><%= L_FAVORITES %></a></span><% end if %></td>
                <td align="middle"><% if m_MainView = "0" then %><a class="menuItem">&nbsp;|&nbsp;</a><% end if %></td>
                <td align="middle"><% if m_MainView = "0" then %><span class="listUnavailable"><%= L_ORGANIZE %></span><% end if %></td>
                <td align="middle"><% if m_MainView = "0" then %><a class="menuItem">&nbsp;|&nbsp;</a><% end if %></td>
                <td align="middle"><a class="menuItem" href="javascript:launchPreferences();"><%= L_PREFERENCES %></a></td>
                <td align="middle"><a class="menuItem">&nbsp;|&nbsp;</a></td>
                <td align="middle"><% if m_MainView = "0" then %><span class="listUnavailable"><%= L_LOGON %></span><% end if %></td>
                <td align="middle"><% if m_MainView = "0" then %><a class="menuItem">&nbsp;|&nbsp;</a><% end if %></td>
                <td align="middle"><a class="menuItem" href="javascript:popHelpWindow();"><%= L_HELP %></a></td></tr></tbody></table></td></tr></tbody></table></td></tr>
  <tr>
    <form name="searchForm" action="available.csp" method="post">
    <td class="menu" noWrap width="64%">
      <table>
        <tbody>
        <tr>
          <td class="menu" noWrap><% if m_MainView = "0" then %><font color="gray"><%= L_LOOKFOR %></font> <input type="hidden" name="otype"> <input type="hidden" value="SI_NAME,SI_ID" name="sortby">
           <input type="hidden" value="personal" name="pageView"> 
           <input class="menuFormElement" contentEditable="false" size="23" name="searchString">
           <select class="menuFormElement" disabled name="searchParameter"> 
           <option value="name" selected><%= L_TITLELCASE %></OPTION>
           <option value="description"><%= L_DESCRIPTIONLCASE %></OPTION>
           <option value="category"><%= L_FOLDERLCASE %></OPTION>
              title<option value="all"><%= L_ALLFIELDSLCASE %></option></select><% end if %></td>
          <td>
            <table cellSpacing="5" cellPadding="0" align="center" border="0">
              <tbody>
              <tr>
                <td class="clsButton">
                  <div class="clsButton"><% if m_MainView = "0" then %><font color="gray"><%= L_SEARCH %></font> <% end if %>
              </div></td></tr></tbody></table></td></tr></tbody></table></td></form>
    <td class="menu" vAlign="top" align="right" width="1%" rowSpan="2"><img alt="<%= L_CORPORATELOGO %>" src="include/menu_dot_bg_default.gif" WIDTH="19" HEIGHT="56"></td>
    <form name="lastfiveform">
    <td class="header" vAlign="center" align="right" width="35%" rowSpan="2"><% if m_MainView = "0" then %><select class="menuFormElement" disabled onchange="javascript:ShowReportLastFive(this.form);" name="lastfive"> 
        <option value="0" selected><%= L_VIEWLAST5 %><option value="0">-<option value="0">-<option value="0">-<option value="0">-<option value="0">-</option></select>&nbsp;<% end if %> </td></form></tr>
  <tr>
    <td class="path"><span class="pathSelected"><% outputMachineName() %></span> 
</td></tr></tbody></table>

<table class="list" cellSpacing="0" cellPadding="3" width="100%" border="0">
  <tbody>
  <tr>
    <td class="list" vAlign="bottom" width="20%"><b><%= L_FOLDERS %></b> <br>
      <hr SIZE="0">
    </td>
    <form>
    <td class="list" vAlign="bottom" noWrap width="80%"><% if m_MainView = "0" then %><b><font color="gray"><%= L_TYPE %></font></b>&nbsp;
    <select class="menuFormElement" style="WIDTH: 180px" disabled onchange="javascript:selectObjectType(this)" name="otype">
     <option value="all" selected><%= L_All %></OPTION>
     <option value="rpt" undefined><%= L_REPORT %></OPTION>
     <option value="arpt" undefined><%= L_ANALYTICALREPORT %></option>
     </select>&nbsp; <b><font color="gray"><%= L_SORTBY %></font></b>&nbsp;
      <select class="menuFormElement" style="WIDTH: 180px" disabled onchange="javascript:doSort(this);" name="sortby"> 
      <option value="SI_NAME,SI_ID" selected><%= L_TITLE %></option>
      <option value="SI_OWNER,SI_NAME,SI_ID"><%= L_OWNER %></option></select> <br>
      <% end if %><hr SIZE="0">
    </td></form></tr>
  <tr><!-- LIST OF SUBCATEGORIES -->
    <td class="list" vAlign="top" width="20%">
    	<% outputCurrentFolder(m_RootItem) %>
		<% outputFolders(m_FolderNodes) %>
	  </td><!-- LIST OF REPORTS -->		
    <td class="list" vAlign="top" width="80%">
      <table cellSpacing="0" cellPadding="0" width="100%" border="0"><!-- Begin list -->
        <tbody>
        <% outputReports(m_ReportNodes) %>

    <td class="list" vAlign="top" align="right"></td>
    <td class="list" vAlign="top" align="right"></td>
  </tr>
  </tbody></table></div>
<script language="JavaScript1.2" src="include/HM_Loader.js" type="text/javascript"></script>
<p>
<table width="100%">
  <tbody>
  <tr><td></td>
</tr></tbody></table>
<tr><td><HR SIZE=0></td><td><HR SIZE=0></td></tr>
<tr>
<td></td>   <td align="right"><span class="list"><%= L_PAGES %> </span>
    <% outputPages m_ReportNodes, m_RootItem  %>
    </td>
</TBODY></TABLE></P>
</DIV>
<DIV align=left>
<a href="http://www.crystaldecisions.com/">
<img alt="<%= L_POWEREDBYCRYSTAL %>" src="include/pb_blue_sml.gif" border="0" WIDTH="108" HEIGHT="30">
</a>
</DIV>
</body>
</html>
