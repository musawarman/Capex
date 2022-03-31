<!--#INCLUDE FILE="defaults.inc" -->
<!--#INCLUDE FILE="LocalizedStrings.asp" -->
<% ' Stateless
' Copyright © 2002 Crystal Decisions, Inc.
const crDirectoryItemTypeFolder	= &H2
const crDirectoryItemTypeReport	= &H4
const crDirectoryItemTypeAllFolders	= &H13
const nMaxReportsDisplay = 9
const nMaxReportsPerRow = 3

' Get the cookie values
Dim m_MainView, m_Viewer

getCookieValues()

Sub getCookieValues()
	' First default to the cookie value	
	m_MainView = Request.Cookies("reportListing")("mainView")
	if (m_MainView = "") then m_MainView = mainViewDefault
	m_Viewer = Request.Cookies("reportListing")("viewer")
	if (m_Viewer = "") then m_Viewer = ViewerDefault
End Sub


Dim ObjectFactory
Set ObjectFactory = CreateObject("CrystalReports.ObjectFactory.2")

' Create ConnectionDirMgr
Dim m_RptAppSession, m_DirMgr
Set m_RptAppSession = ObjectFactory.CreateObject("CrystalReports.ReportAppSession")
m_RptAppSession.Initialize
Set m_DirMgr = m_RptAppSession.CreateService("CrystalReports.ConnectionDirManager.2")
m_DirMgr.Open crDirectoryItemTypeAllFolders or crDirectoryItemTypeReport

Dim m_Folder, m_Nodes, m_RootItem
m_Folder = Request.QueryString.Item("folder")
Set m_RootItem = ObjectFactory.CreateObject("CrystalReports.ConnectionDirectoryItem")
Set m_Nodes = getReportsAndFolders(m_Folder)

Dim m_ReportNodes, m_FolderNodes
Set m_ReportNodes = ObjectFactory.CreateObject("CrystalReports.DirectoryItems")
Set m_FolderNodes = ObjectFactory.CreateObject("CrystalReports.DirectoryItems")
separateNodeTypes m_Nodes, m_ReportNodes, m_FolderNodes

' Iterate through and separate the reports and folders
Sub separateNodeTypes(allNodes, reportNodes, folderNodes)
	Dim i, item
	for i = 0 to allNodes.count - 1
		Set item = allNodes.item(i)
		if item.isLeaf then
			reportNodes.add item
		else
			folderNodes.add item
		end if
	next
End Sub

Dim m_nStartIndex
m_nStartIndex = Request.QueryString.Item("startIdx")
if m_nStartIndex = "" then
	m_nStartIndex = 0
else
	m_nStartIndex = CLng(m_nStartIndex)
end if

Function getReportsAndFolders(parentID)
	Dim dirItems
	if parentID = "" then
		Set dirItems = m_DirMgr.getRoots()
		Set m_RootItem = nothing  
		Set dirItems = m_DirMgr.getChildren(dirItems.Item(0))
	else 
		m_RootItem.UID.FromString(parentID)
		m_RootItem.Name = Request.QueryString.Item("Name")
		Set dirItems = m_DirMgr.getChildren(m_RootItem)
	end if
	Set getReportsAndFolders = dirItems
End Function

Function outputCurrentFolder(item)
	if (item is Nothing) then Exit Function
	Response.Write "<a class='list' href='reportListing.asp?folder=" + Server.URLEncode(item.UID.ToString)+ "&name=" + Server.URLEncode(item.Name) + "'>"
	Response.Write Server.HTMLEncode(item.Name) + "</a>"  + vbCrLf
End Function

' Assume dirItems are only folders
Function outputFolders(dirItems)
	Dim i, item
	For i = 0 to dirItems.count - 1
		Set item = dirItems.Item(i)
		Response.Write "<li>"
		outputCurrentFolder(item)
		Response.Write "<br></li>" + vbCrLf
	next
End Function

' Assume dirItems are only reports
Function outputReports(dirItems)
	Dim i, item, stopIndex
	if (m_nStartIndex + nMaxReportsDisplay < dirItems.count) then
		stopIndex = m_nStartIndex + nMaxReportsDisplay
	else 
		stopIndex = dirItems.count
	end if
	
	For i = m_nStartIndex to stopIndex - 1
		if (i mod nMaxReportsPerRow = 0) then Response.Write "<tr>" + vbCrLf
		Set item = dirItems.Item(i)

		' Generate the correct link for the viewer
		Dim sViewerLink
		if (m_Viewer = "0" or m_Viewer = "3") then ' ActiveX Viewer or Java Viewer
			sViewerLink = "href='rptServer/Start.asp?ReportName=" + Server.URLEncode(item.UID.StringValue("path")) + "&Viewer=" + m_Viewer + "'"
		elseif m_Viewer = "1" then ' Page Viewer	
			sViewerLink = "href='HTMLViewers/pageViewer.asp?ReportName=" + Server.URLEncode(item.UID.StringValue("path")) + "'"
		elseif m_Viewer = "4" then ' Parts Viewer
			sViewerLink = "href='HTMLViewers/partsViewer.asp?ReportName=" + Server.URLEncode(item.UID.StringValue("path")) + "'"
		else	' Default to interactive viewer
			sViewerLink = "href='HTMLViewers/interactiveViewer.asp?ReportName=" + Server.URLEncode(item.UID.StringValue("path")) + "'"
		end if			

		Response.Write "<td vAlign='top' width='34%'><table><tbody><tr>" + vbCrLf
		Response.Write "  <td vAlign='top'><a class='list' " + sViewerLink + " TARGET='_blank'>" + vbCrLf
		Response.Write "  <img class='list' src='include/thumbnail.gif' width='40' align='left' border='1'></a></td>" + vbCrLf
		Response.Write "  <td vAlign='top'><div class='list'>" + Server.HTMLEncode(item.Name) + "<br>" + vbCrLf
		Response.Write "   <a class='list' " + sViewerLink + " TARGET='_blank'>"
		
		Response.Write L_VIEW + "</a>"
		' Add greyed out links for Schedule and History
		if m_MainView = "0" then Response.Write " <span class='itemUnavailable'> | " + L_SCHEDULE + " | " + L_HISTORY + " </span>" + vbCrLf
		Response.Write "</div></td></tr></tbody></table></td>" + vbCrLf
		
		if ((i mod nMaxReportsPerRow = nMaxReportsPerRow - 1) and (i <> stopIndex - 1)) then _
			Response.Write "<tr><td colSpan='" + CStr(nMaxReportsPerRow) + "'><hr SIZE='0'></td></tr>" + vbCrLf
	next
End Function

' Assume dirItems are only reports
Sub outputPages(dirItems, currentFolder)
	Dim i, nMaxIdx
	nMaxIdx = dirItems.Count / nMaxReportsDisplay	' zero indexed
	
	Dim folderID, folderName
	if (currentFolder is nothing) then
		folderID = ""
		folderName = ""
	else 
		folderID = Server.URLEncode(currentFolder.UID.ToString)
		folderName = Server.URLEncode(currentFolder.Name)
	end if
	
	for i = 0 to nMaxIdx
		if (i = m_nStartIndex / nMaxReportsDisplay) then
			Response.Write "<span class='listSelected'>[" + CStr(i + 1) + "]</span>"
		else
			Response.Write "<a class='list' href='reportListing.asp?startIdx=" + _
						    CStr(i * nMaxReportsDisplay) +"&folder=" + _
							 folderID + "&name=" + folderName + _
							 "'>[" + CStr(i + 1) + "]</a>"
		end if
	next
End Sub


Sub outputMachineName()
	Response.Write "<a class='menuItem' href=reportListing.asp>" + UCase(Request.ServerVariables.Item("SERVER_NAME")) + "</a>" + vbCrLf
End Sub

%>


