
function OpenWindow(url, name) {
	window.open(url, name);
}

function ShowPopup(url, name, width, height, scrollable, resizable) {
	window.open(url, name, "width=" + width + ", height=" + height + ", resizable=" + (resizable ? "1" : "0") + ", scrollbars=" + (scrollable ? "1" : "0") + ", status=1");
}

function ShowHtmlPopup(url, title, inIframe) {
	
	$("#htmlPopupContainer").show();
	$("#htmlPopupContainer").css("height", $(document).height());

	$("#htmlPopup").css("top", document.documentElement.scrollTop);

	$("#htmPopupTitleContent").html(unescape(title));

	if (inIframe)
		$("#htmlPopupContent").html("<iframe id=\"popupContentIFrame\" name=\"popupContentIFrame\" src=\"" + url + "\" width=\"100%\" height=\"500\" frameborder=\"0\" marginheight=\"0\" marginwidth=\"0\" scrolling=\"auto\" />");
	else
		$("#htmlPopupContent").html("<iframe id=\"popupContentIFrame\" name=\"popupContentIFrame\" src=\"" + url + "\" width=\"100%\" height=\"500\" frameborder=\"0\" marginheight=\"0\" marginwidth=\"0\" scrolling=\"auto\" />");
	//	$("#htmlPopupContent").load(url);

	//	$.ajax({ url: "/Intranet", success: function (result) { popupContent.innerHTML = result; } });
}

function HideHtmlPopup() {
	$("#htmlPopupContainer").hide();
	$("#htmPopupTitleContent").html("");
	$("#htmlPopupContent").html("");
}


function ShowHideElement(elementId) {
	var element = document.getElementById(elementId);
	if (element == null) return;

	if (element.style.display == "none")
		element.style.display = "block";
	else
		element.style.display = "none";
}

function ShowHideElementWIcon(elementId, iconObj, iconSrcTemplate) {
	var element = document.getElementById(elementId);
	if (element == null) return;

	if (element.style.display == "none")
		element.style.display = "block";
	else
		element.style.display = "none";

	if (iconObj == null || iconSrcTemplate == "") return;
	// sync the icon to the element's current condition:
	if (element.style.display == "none") 
		iconObj.src = iconSrcTemplate.replace("{0}", "closed");
	else
		iconObj.src = iconSrcTemplate.replace("{0}", "opened");
}

function ShowHideElementWLink(elementId, linkObj, linkApplyOnStyle) {
	var element = document.getElementById(elementId);
	if (element == null) return;

	if (element.style.display == "none")
		element.style.display = "block";
	else
		element.style.display = "none";

	if (linkObj == null || linkApplyOnStyle == "") return;
	// sync the link to the element's current condition:
	if (element.style.display == "none")
		linkObj.className = linkObj.className.replace(" " + linkApplyOnStyle, "");
	else
		linkObj.className = linkObj.className + " " + linkApplyOnStyle;
}

function ShowHideTab(tabObj, tabContainerId) {
	var allTabs = $("#" + tabContainerId + " a");
	$("#" + tabContainerId + " a").attr("class", "");
	for (i = 0; i < allTabs.length; i++) {
		if (tabObj.id != allTabs[i].id) {
			$("#" + allTabs[i].id + "_content").hide();
		}
		else {
			allTabs[i].className = "active";
			$("#" + allTabs[i].id + "_content").show();
		}
	}
}


function LoadDropdownList(ddlId) {
	var svc = new ICAAjaxServices.Common();

	if (ddlId == "project_statuses")
		svc.GetProjectStatuses(ddlId, LoadDropdownList_OnSuccess, null, urlParams.Status);

	else if (ddlId == "project_countries")
		svc.GetProjectCountries(ddlId, LoadDropdownList_OnSuccess, null, urlParams.Country);

	else if (ddlId == "vacancy_countries")
		svc.GetRegions(ddlId, LoadDropdownList_OnSuccess, null, urlParams.Region);

	else if (ddlId == "vacancy_sectors")
		svc.GetMainSectorsForVacancySearch(ddlId, LoadDropdownList_OnSuccess, null, urlParams.MainSector);

	else if (ddlId == "project_fundingAgencies")
		svc.GetProjectFundingAgenciesAndIDs(ddlId, LoadDropdownList_OnSuccess, null, urlParams.Donor);

	else if (ddlId == "reference_fundingAgencies")
		svc.GetProjectFundingAgencies(ddlId, LoadDropdownList_OnSuccess, null, urlParams.Donor);

	else if (ddlId == "period_from")
		svc.GetReferencesFilterYearFrom(ddlId, LoadDropdownList_OnSuccess, null, urlParams.YearFrom);

	else if (ddlId == "period_to")
		svc.GetReferencesFilterYearTo(ddlId, LoadDropdownList_OnSuccess, null, urlParams.YearTo);
}

function LoadDropdownList_OnSuccess(result, userCtx) {

	var ddlOpt = document.getElementById(result.ClientId).options;

	for (i = 0; i < result.Items.length; i++) {
		ddlOpt[ddlOpt.length] = new Option(result.Items[i].Text, result.Items[i].Value);
		if (userCtx != null && userCtx != '') {
			if (userCtx == result.Items[i].Value)
				document.getElementById(result.ClientId).selectedIndex = ddlOpt.length;
		}
	}
}


function SubmitForm(formId, actionUrl, target) {
	var form = formId != null && formId != "" ? $("#" + formId) : document.forms[0];
	if (form != null) {
		if (actionUrl != null && actionUrl != "") form.action = actionUrl;
		if (target != null && target != "") form.target = target;

		form.submit();
	}
}

function IsNullOrEmpty(str) {
	return str == null || str == "";
}

function ajax_load(path, containerId, resizeParent) {
	ajax_load_eval(path, containerId, resizeParent, "");
	return false;
}

function ajax_load_eval(path, containerId, resizeParent, evalCode) {
	$.ajax({
		url: path,
		success: function (result) {
			$('#' + containerId).html(result);
			if (resizeParent) ResizeParentFrame();
			if (evalCode != "")
				eval(evalCode);
		}
	});
	return false;
}

function ajax_post(formId, path, evalOnSuccess) {

	$.ajax({
		url: path,
		type: "POST",
		enctype: "multipart/form-data", 
		data: $("#" + formId).serialize(),
		success: function (result) { eval(evalOnSuccess); }
	});

	return false;
}

function HideMessages() {
	$("#success").html('');
	$("#error").html('');
}

function ShowMessage(successMsg, errorMsg) {
	$("#success").html(successMsg + '<br/>');
	$("#error").html(errorMsg + '<br/>');
}

function AppendMessage(successMsg, errorMsg) {
	$("#success").html($("#success").html() + successMsg + '<br/>');
	$("#error").html($("#error").html() + errorMsg + '<br/>');
}

function ResizeParentFrame() {
	if ($(document).height() > 0) 
		parent.$('#popupContentIFrame').css('height', $(document).height());
}

function PrjForm_AddReference_Click() {
	ajax_load_eval('/Intranet/EditReference', 'projReferences_form', true, "PrjForm_AddReference_CopyProjectInfo();");
	$("#addProjReferenceBtn").hide();
}

function PrjForm_EditReference_Click(sender, referenceId, project2Id, languageId) {
	//alert('/Intranet/EditReference?id=' + referenceId + '&p2id=' + project2Id + '&lid=' + languageId);
	if (languageId != '' && (project2Id != '' || referenceId != ''))
		ajax_load_eval('/Intranet/EditReference?id=' + referenceId + '&p2id=' + project2Id + '&lid=' + languageId, 'projReferences_form', true, "PrjForm_AddReference_CopyProjectInfo();");

//	$("#referencesForProjectList tr").css("background", "");
//	sender.style.background = "#EEEEEE";
//	$("#addProjReferenceBtn").hide();
}

function PrjForm_EditReference_AfterFinished(projectId) {
	// update references list:
//	ajax_load('/Intranet/ProjectReferencesList?pid=' + projectId, 'projReferences_existing', true);
//	$("#projReferences_form").html('');
	// show 'add reference' button:
//	$("#addProjReferenceBtn").show();
}

function PrjForm_AddReference_CopyProjectInfo() {

	if ($("#Ref_Title").val() == '') $("#Ref_Title").val($("#projectTitle").val());

	if ($("#Ref_Country").val() == '') {
		var prjCountries = $("#divCountriesDisplayListOnly").html();
		var i = 0;
		debugger;
		while ((prjCountries.toLowerCase().indexOf("<br/>") > -1 || prjCountries.toLowerCase().indexOf("<br>") > -1 > -1) && i < 50) {
			prjCountries = prjCountries.replace("\n", ", ").replace("<BR/>", ", ").replace("<BR>", ", ").replace("<br/>", ", ").replace("<br>", ", ");
			//avoid endless loop:
			i++;
		}
		$("#Ref_Country").val(prjCountries);
	}
	
	if ($("#Ref_BudgetAmount").val() == '') $("#Ref_BudgetAmount").val($("#budget").val());
	if ($("#Ref_IdCurrency").val() == '') $("#Ref_IdCurrency").val($("#idCurrency").val());
	if ($("#fldRefFunding").val() == '') $("#fldRefFunding").val($("#fldDonorName").val());
}

function HelpIcon_MouseClick(selfObj) {
	var helpObj = document.getElementById("htmlHelpContent");
	if (helpObj != null) {
		if (helpObj.style.display == "block") {
			HelpIcon_MouseOut();
		}
		else {
			HelpIcon_MouseOver(selfObj);
		}
	}
}

function HelpIcon_MouseOver(selfObj) {
	var helpObj = document.getElementById("htmlHelpContent");
	if (helpObj != null) {
		helpObj.innerHTML = unescape(selfObj.helptext);
		helpObj.style.display = "block";
		SetHelpPosition(helpObj);
	}
}

function HelpIcon_MouseOut() {
	var helpObj = document.getElementById("htmlHelpContent");
	if (helpObj != null) {
		helpObj.innerHTML = "";
		helpObj.style.display = "none";
	}
}

function SetHelpPosition(helpObj) {
	var posTop = window.event.clientY - helpObj.scrollHeight - 10; // document.body.scrollTop
	if (posTop < 0) posTop = 0;
	helpObj.style.top = posTop;

	var posLeft = window.event.clientX - helpObj.scrollWidth - 10;
	if (posLeft < 0) posLeft = 0;
	helpObj.style.left = posLeft;
}




function PrjForm_AddPosition_Click() {
	ajax_load_eval('/Intranet/EditPosition', 'projPositions_form', true, "");
	$("#addProjPositionBtn").hide();
}

function PrjForm_EditPosition_Click(sender, positionId) {
	ajax_load('/Intranet/EditPosition?id=' + positionId, 'projPositions_form', true);
	$("#positionsForProjectList tr").css("background", "");
	sender.style.background = "#EEEEEE";
	$("#addProjPositionBtn").hide();
}

function PrjForm_EditPosition_AfterFinished(project2Id) {
	// update positions list:
	ajax_load('/Intranet/ProjectPositionsList?p2id=' + project2Id, 'projPositions_existing', true);
	$("#projPositions_form").html('');
	// show 'add position' button:
	$("#addProjPositionBtn").show();
}

function DeletePosition(posId) {
	if (posId > 0) {
		var svc = new ICAAjaxServices.Positions();
		svc.DeletePosition(posId, DeletePosition_OnSuccess, null, null);
	}
}

function DeletePosition_OnSuccess(result) {
	if (result) {
		// reload the positions list:
		eval(DeletePosition_OnSuccessAction);
	}
	else
		alert("not deleted");
}

function AddPartnerCompany() {
	if (CurrentProject2Id && CurrentProject2Id > 0) {
		var svc = new ICAAjaxServices.Common();
		var result = svc.AddProjectPartnerCompany(
			$("#fldCompany").val(),
			CurrentProject2Id,
			$("#roleId").val(), 
			$("#fldComment").val(), EditPartnerCompany_OnSuccess, null, null);
	}
	else
		alert("Unable to add partner company.");
}

function DeletePartnerCompany(companyId) {
	if (CurrentProject2Id && CurrentProject2Id > 0) {
		var svc = new ICAAjaxServices.Common();
		svc.RemoveProjectPartnerCompany(companyId, CurrentProject2Id, EditPartnerCompany_OnSuccess, null, null);
	}
	else
		alert("Unable to delete partner company.");
}

function EditPartnerCompany_OnSuccess(result) {
	if (CurrentProject2Id && CurrentProject2Id > 0)
		ajax_load('/Intranet/ProjectPartners?p2id=' + CurrentProject2Id, 'projPartners_content', true);
	else
		alert("Unable to reload partner companies list.");
}


function FormatNumber(number, thousandsSep, decimalsCount, decimalSep) {
	
	if (number % 1 == 0) decimalsCount = 0;
	debugger;
	var tmp_result = number.toFixed(decimalsCount) + '';
	var resSpl = tmp_result.split('.');
	var result = "";
	var index = 0;
	if (resSpl[0].length > 0) {
		for (ifn = resSpl[0].length - 1; ifn >= 0; ifn--) {

			result = resSpl[0][ifn] + '' + result;
			index++;

			if (index > 2) {
				index = 0;
				if (ifn > 0)
					result = thousandsSep + result;
			}
		}

		if (resSpl.length > 1) {
			result = result + decimalSep + resSpl[1];
		}
	}

	return result;
}

function ClearInputValue(inputId) {
	if (document.getElementById(inputId)) {
		document.getElementById(inputId).value = "";
	}
	return;
}

function ResetFilter() {
	if (FormClientId != '') {
		document.getElementById(FormClientId).reset();
		$("#divSectorsDisplayList").html("--- All ---");
		$("#hidSectors").val("");
	}
}


/* changed after 25.05: */


/* new after 25.05: */

function BuildHtmlListItem(containerId, itemId, itemTitle, hidIDsFldId) {
	var template = "{0}<div id=\"lstObj_{4}\" class=\"lstItem\"><div class=\"itr\"><a href=\"javascript:RemoveHtmlListItem('lstObj_{5}', {1}, '{3}')\" title=\"Remove item\">&nbsp;</a></div><span class=\"itxt\">{2}</span></div>";
	template = template.replace('{0}', $('#' + containerId).html());
	template = template.replace('{1}', itemId);
	template = template.replace('{2}', itemTitle);
	template = template.replace('{3}', hidIDsFldId);
	template = template.replace('{4}', itemId);
	template = template.replace('{5}', itemId);

	$('#' + containerId).html(template);
}

function RemoveHtmlListItem(listObj, idItem, hidIDsFldId) {
	if (hidIDsFldId.toLowerCase() == "hidsectors") {
		RemoveSector(idItem, hidIDsFldId);
	}
	else if (hidIDsFldId.toLowerCase() == "hidcountries") {
		RemoveCountry(idItem, hidIDsFldId);
	}
	else if (hidIDsFldId.toLowerCase() == "hidpartners") {
		DeletePartnerCompany(idItem);
	}

	$('#' + listObj).hide();
}

function RemoveSector(id, hidIDsFldId) {
	selectedSectorIDs = $('#' + hidIDsFldId).val();
	if (selectedSectorIDs.indexOf(id) > -1) {
		var removeSectorStr = id;
		// check if the id is first in the list and there are more behind it:
		if (selectedSectorIDs.indexOf(id + ",") > -1) removeSectorStr = removeSectorStr + ",";
		// check if there's a comma in front of the id:
		else if (selectedSectorIDs.indexOf("," + id) > -1) removeSectorStr = "," + removeSectorStr;

		selectedSectorIDs = selectedSectorIDs.replace(removeSectorStr, "");

		$('#' + hidIDsFldId).val(selectedSectorIDs);
	}
}

function RemoveCountry(id, hidIDsFldId) {
	selectedCountryIDs = $('#' + hidIDsFldId).val();
	if (selectedCountryIDs.indexOf(id) > -1) {
		var removeCountryStr = id;
		// check if the id is first in the list and there are more behind it:
		if (selectedCountryIDs.indexOf(id + ",") > -1) removeCountryStr = removeCountryStr + ",";
		// check if there's a comma in front of the id:
		else if (selectedCountryIDs.indexOf("," + id) > -1) removeCountryStr = "," + removeCountryStr;

		selectedCountryIDs = selectedCountryIDs.replace(removeCountryStr, "");

		$('#' + hidIDsFldId).val(selectedCountryIDs);
	}
}



/* from new design: */

function toggle(div_id) {
	var el = document.getElementById(div_id);
	if (el.style.display == 'none') { el.style.display = 'block'; }
	else { el.style.display = 'none'; }
}
function blanket_size(popUpDivVar) {
	if (typeof window.innerWidth != 'undefined') {
		viewportheight = window.innerHeight;
	} else {
		viewportheight = document.documentElement.clientHeight;
	}
	if ((viewportheight > document.body.parentNode.scrollHeight) && (viewportheight > document.body.parentNode.clientHeight)) {
		blanket_height = viewportheight;
	} else {
		if (document.body.parentNode.clientHeight > document.body.parentNode.scrollHeight) {
			blanket_height = document.body.parentNode.clientHeight;
		} else {
			blanket_height = document.body.parentNode.scrollHeight;
		}
	}
	var blanket = document.getElementById('blanket');
	blanket.style.height = blanket_height + 'px';
	var popUpDiv = document.getElementById(popUpDivVar);
	popUpDiv_height = blanket_height / 2 - 150; //150 is half popup's height
	popUpDiv.style.top = popUpDiv_height + 'px';
}
function window_pos(popUpDivVar) {
	if (typeof window.innerWidth != 'undefined') {
		viewportwidth = window.innerHeight;
	} else {
		viewportwidth = document.documentElement.clientHeight;
	}
	if ((viewportwidth > document.body.parentNode.scrollWidth) && (viewportwidth > document.body.parentNode.clientWidth)) {
		window_width = viewportwidth;
	} else {
		if (document.body.parentNode.clientWidth > document.body.parentNode.scrollWidth) {
			window_width = document.body.parentNode.clientWidth;
		} else {
			window_width = document.body.parentNode.scrollWidth;
		}
	}
	var popUpDiv = document.getElementById(popUpDivVar);
	window_width = window_width / 2 - 150; //150 is half popup's width
	popUpDiv.style.left = window_width + 'px';
}
function popup(windowname) {
	blanket_size(windowname);
	window_pos(windowname);
	toggle('blanket');
	toggle(windowname);
}
/* end of - from new design */