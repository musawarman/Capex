// buttons.js
//
// This file contains the functions and array of images for changing the images on the 
// menu buttons when the user rolls their mouse over or clicks on a button.
//

var imgArr = new Array();

imgArr["butHelp"] = link_path + "images/help_up_";
imgArr["butHelpOver"] = link_path + "images/help_over_";
imgArr["butHelpDown"] = link_path + "images/help_down_";

imgArr["butClose"] = link_path + "images/close_up_";
imgArr["butCloseOver"] = link_path + "images/close_over_";
imgArr["butCloseDown"] = link_path + "images/close_down_";

imgArr["butExit"] = link_path + "images/exit_up_";
imgArr["butExitOver"] = link_path + "images/exit_over_";
imgArr["butExitDown"] = link_path + "images/exit_down_";

imgArr["butLogon"] = link_path + "images/logon_up_";
imgArr["butLogonOver"] = link_path + "images/logon_over_";
imgArr["butLogonDown"] = link_path + "images/logon_down_";

imgArr["butSignup"] = link_path + "images/signup_up_";
imgArr["butSignupOver"] = link_path + "images/signup_over_";
imgArr["butSignupDown"] = link_path + "images/signup_down_";

imgArr["butBack"] = link_path + "images/back_up_";
imgArr["butBackOver"] = link_path + "images/back_over_";
imgArr["butBackDown"] = link_path + "images/back_down_";

imgArr["butAlerts"] = link_path + "images/alerts_up_";
imgArr["butAlertsOver"] = link_path + "images/alerts_over_";
imgArr["butAlertsDown"] = link_path + "images/alerts_down_";

imgArr["butLogoff"] = link_path + "images/logoff_up_";
imgArr["butLogoffOver"] = link_path + "images/logoff_over_";
imgArr["butLogoffDown"] = link_path + "images/logoff_down_";

imgArr["butSettings"] = link_path + "images/preferences_up_";
imgArr["butSettingsOver"] = link_path + "images/preferences_over_";
imgArr["butSettingsDown"] = link_path + "images/preferences_down_";

imgArr["butOrganize"] = link_path + "images/organize_up_";
imgArr["butOrganizeOver"] = link_path + "images/organize_over_";
imgArr["butOrganizeDown"] = link_path + "images/organize_down_";

imgArr["butFavorites"] = link_path + "images/favorites_up_";
imgArr["butFavoritesOver"] = link_path + "images/favorites_over_";
imgArr["butFavoritesDown"] = link_path + "images/favorites_down_";

imgArr["butPublic"] = link_path + "images/public_up_";
imgArr["butPublicOver"] = link_path + "images/public_over_";
imgArr["butPublicDown"] = link_path + "images/public_down_";

imgArr["butApply"] = link_path + "images/apply_up_";
imgArr["butApplyOver"] = link_path + "images/apply_over_";
imgArr["butApplyDown"] = link_path + "images/apply_down_";

imgArr["butRefresh"] = link_path + "images/refresh_up_";
imgArr["butRefreshOver"] = link_path + "images/refresh_over_";
imgArr["butRefreshDown"] = link_path + "images/refresh_down_";

imgArr["butSchedule"] = link_path + "images/schedule_up_";
imgArr["butScheduleOver"] = link_path + "images/schedule_over_";
imgArr["butScheduleDown"] = link_path + "images/schedule_down_";

imgArr["butPassword"] = link_path + "images/password_up_";
imgArr["butPasswordOver"] = link_path + "images/password_over_";
imgArr["butPasswordDown"] = link_path + "images/password_down_";

function linkOver(linkName, style) {
	if (document.images)
		document[linkName].src = imgArr['but' + linkName + 'Over'] + style + ".gif";
}

function linkDown(linkName, style) {
	if (document.images) {
		document[linkName].src = imgArr['but' + linkName + 'Down'] + style + ".gif";
	}
}

function linkNormal(linkName, style) {
	if (document.images)
		document[linkName].src = imgArr['but' + linkName] + style + ".gif";
}

