/*
	File Version Start - Do not remove this if you are modifying the file 
	Build: 8.5.2
	File Version End
*/

// pophelpwindow.js
//
// This file contains a function that will open a new window to display the ePortfolio help.
//

function popHelpWindow() {

	var sw = 800;  //SCREEN WIDTH VARIABLE
	var sh = 600;  //SCREEN HEIGHT VARIABLE

	//CENTER WINDOW ON SCREEN
	LeftPosition = (screen.width) ? (screen.width-sw)/2 : 0;
	TopPosition = (screen.height) ? (screen.height-sh)/2 : 0;

	winProps = "width="+sw+",height="+sh+",location=no,scrollbars=yes,menubar=no,toolbar=no,resizable=yes,top="+TopPosition+",left="+LeftPosition;
	helpWindow = window.open("http://www.crystaldecisions.com/ipl/default.asp?destination=webreporting&product=crystalreports&language=en", "", winProps);
}
