
var AddListPaging = true;

// PageTypeName possible values:
// MYTB
// MYTB-CS
// ICATB
// ICATB-CS
var PageTypeName = "ICATB"; // default

function LoadProjectList(pageSize, pageNumber, containerId) {
	var svc = new ICAAjaxServices.Projects();
	
	svc.GetProjectList(
		urlParams.ProjectType,
		urlParams.Keyword,
		urlParams.DonorId,
		urlParams.Country, 
		urlParams.SectorList,
		urlParams.Status,
		urlParams.Unit,
		urlParams.BudgetMin,
		urlParams.BudgetMax,
		urlParams.Orderby, 
		urlParams.CompanyId,
		urlParams.GetCompanyProjectsOnly,
		urlParams.IncludeNonAwarded,
		pageSize,
		pageNumber,
		containerId, 
		LoadProjectList_OnSuccess, null, null);
}

function LoadProjectList_OnSuccess(result) {

	// release lock:
	LockListExecution = 0;

	// mark the list as "finished", to not be appended anymore:
	if (result.ProjectList.length < result.PageSize) {
		itemFinishedArray[result.ContainerId] = 1;
	}

	if ((result.ProjectList == null || result.ProjectList.length == 0) && $("#" + result.ContainerId + " tr").size() < 2) {
		List_AddNoResultsRow(result.ContainerId);
		return;
	}

	for (i = 0; i < result.ProjectList.length; i++) {
		ProjectList_AddItemRow(result.ContainerId, result.ProjectList[i]);
	}

	if (AddListPaging)
		List_AddPagingRow(result.ContainerId, result.CurrentPage, Math.ceil(result.ProjectsTotalCount / result.PageSize));

}

function List_Update(tblId, resetPageCounter) {
	// delay execution if locked:
	if (LockListExecution != 0) {
		setTimeout("List_Update('" + tblId + "', " + resetPageCounter + ")", 500);
		return;
	}

	// lock execution:
	LockListExecution = 1;

	if (StatusIdsTB[tblId] != null) // to avoid overwriting the status for the custom status search with the filter
		urlParams.Status = StatusIdsTB[tblId];

	// include 'not awarded', because there are such statuses:
	if (urlParams.Status >= 1 && urlParams.Status < 20
		|| urlParams.Status > 24 && urlParams.Status != 30 && urlParams.Status != 31 && urlParams.Status != 40) {
		urlParams.IncludeNonAwarded = true;
	}

	// clear table:
	$("#" + tblId + " tr:gt(0)").remove();

	// clear count:
	$("#" + tblId.replace("tbl", "count")).html("");

	// set custom results headers:
	SetCustomStatusResultsHeaders(urlParams.Status);

	if (resetPageCounter)
		List_ResetPageCounter();

	LoadProjectList(ListPageSize, ListCurrentPageNumber, tblId);
}

function List_Append(tblId) {
	// delay execution if locked:
	if (LockListExecution != 0) {
		setTimeout("List_Append('" + tblId + "')", 500);
		return;
	}

	// lock execution:
	LockListExecution = 1;

	if (StatusIdsTB[tblId] != null) // to avoid overwriting the status for the custom status search with the filter
		urlParams.Status = StatusIdsTB[tblId];

	// include 'not awarded', because there are such statuses:
	if (urlParams.Status >= 1 && urlParams.Status < 20
		|| urlParams.Status > 24 && urlParams.Status != 30 && urlParams.Status != 31 && urlParams.Status != 40) {
		urlParams.IncludeNonAwarded = true;
	}

	LoadProjectList(ListPageSize, ListCurrentPageNumber, tblId);
}

function List_Update_Tenderboard(tblId, resetPageCounter, projectStatus) {
	// delay execution if locked:
	if (LockListExecution != 0) {
		setTimeout("List_Update_Tenderboard('" + tblId + "', " + resetPageCounter + ", " + projectStatus + ")", 500);
		return;
	}

	// lock execution:
	LockListExecution = 1;

	if (StatusIdsTB[tblId] != null) // to avoid overwriting the status for the custom status search with the filter
		urlParams.Status = StatusIdsTB[tblId];

	// include 'not awarded', because there are such statuses:
	urlParams.IncludeNonAwarded = true;
	
	// clear table:
	$("#" + tblId + " tr:gt(0)").remove();

	// clear count:
	$("#" + tblId.replace("tbl", "count")).html("");

	if (resetPageCounter)
		List_ResetPageCounter();

	LoadProjectList(ListPageSize, ListCurrentPageNumber, tblId);
}

function CheckPositionAndLoad(item, projectStatus) {

	if (LockListExecution != 0) {
		setTimeout("CheckPositionAndLoad('" + item + "', " + projectStatus + ")", 500);
		return;
	}

	var loadBorder = document.documentElement.scrollTop + $(window).height() + 300;
	if (loadBorder >= $('#tbl' + item + ' tr:last').position().top) {

		// check if current item was added in itemPageArray:
		var foundInPageArray = false;
		for (var pageKey in itemPageArray) {
			if (pageKey == item) {
				foundInPageArray = true;
				break;
			}
		}

		var alreadyFullyLoaded = false;
		for (var itemKey in itemFinishedArray) {
			if (itemKey == "tbl" + item) {
				alreadyFullyLoaded = true;
				break;
			}
		}

		// if still not fully loaded:
		if (!alreadyFullyLoaded) {
			// if no rows added - load the list:
			if (!foundInPageArray) {
				if (projectStatus > 0)
					List_Update_Tenderboard("tbl" + item, true, projectStatus);
				else
					List_Update("tbl" + item, true);
				itemPageArray[item] = 1;
			}
			else {
				// if rows added - append next page to the result:
				itemPageArray[item]++;
				List_AppendPage(item, itemPageArray[item]);
			}
		}
	}
}


function GetParamsFromSearchForm() {
	urlParams.ProjectType = $("#project_type") != null && $("#project_type").children("option:selected").length > 0 ? $("#project_type").children("option:selected")[0].value : null;
	if (urlParams.ProjectType == '') urlParams.ProjectType = null;
	urlParams.Keyword = $("#keywords").attr("value");
	urlParams.DonorId = $("#project_fundingAgencies") != null && $("#project_fundingAgencies").children("option:selected").length > 0 ? $("#project_fundingAgencies").children("option:selected")[0].value : null;
	if (urlParams.DonorId == '') urlParams.DonorId = null;
	urlParams.Country = $("#project_countries") != null && $("#project_countries").children("option:selected").length > 0 ? $("#project_countries").children("option:selected")[0].value : "";
	urlParams.SectorList = $("#hidSectors").attr("value") != "" ? $("#hidSectors").attr("value") : null;
	urlParams.Status = $("#project_statuses")[0] != null && $("#project_statuses").children("option:selected").length > 0 && $("#project_statuses").children("option:selected")[0].value != '' ? $("#project_statuses").children("option:selected")[0].value : null;
	urlParams.Unit = "";
	urlParams.BudgetMin = parseInt($("#budget_from").attr("value")) > 0 ? parseInt($("#budget_from").attr("value")) : null;
	urlParams.BudgetMax = parseInt($("#budget_to").attr("value")) > 0 ? parseInt($("#budget_to").attr("value")) : null;

	// check it by default if a 'non-awarded' status selected (awarded statuses: 20, 21, 22, 23, 24, 30, 31, 40):
	if (urlParams.Status >= 1 && urlParams.Status < 20 
		|| urlParams.Status > 24 && urlParams.Status != 30 && urlParams.Status != 31 && urlParams.Status != 40) {
		document.getElementById("cbIncludeNotAwarded").checked = true;
	}
	urlParams.IncludeNonAwarded = document.getElementById("cbIncludeNotAwarded").checked;

	urlParams.Orderby = "";
}


function ProjectList_AddItemRow(tableId, obj) {

	if (obj.Title == null) return;
	if (obj.Title == "") return;

	var table = document.getElementById(tableId);
	if (table == null) return;

	var row = table.insertRow(table.rows.length);
	row.id = "row_" + obj.ID;
	row.valign = "top";
	row.className = "degree_" + obj.ImportanceValue;
	row.onmouseover = function () { this.style.backgroundColor = '#f3f3f3'; };
	row.onmouseout = function () { this.style.backgroundColor = '#fff'; };

	var cellId = 0;

	// Countries:
	var cell0 = row.insertCell(cellId++);
	cell0.innerHTML = obj.Countries != null ? obj.Countries : "&nbsp;";
	
	// Funding agency (donor):
	var cell1 = row.insertCell(cellId++);
	cell1.innerHTML = obj.DonorName != null ? obj.DonorName : "&nbsp;";

	// Title:
	var cell2 = row.insertCell(cellId++);
	if (PageTypeName == "MYTB" || PageTypeName == "MYTB-CS") {
		cell2.innerHTML = "";
		cell2.innerHTML += "<span id=\"p_" + obj.ID + "\">" +
		"<a href=\"javascript: void(0)\" onclick=\"ShowHtmlPopup('" + EditProjectUrl + "?id=" + obj.ID + "', 'Edit Project: " + escape(obj.Title) + "', true)\">" +
		(obj.Title != null ? obj.Title : "<i>&laquo; N/A &raquo;</i>") + "</a><i" + (obj.StatusUpdate != null && obj.StatusUpdate == "1" ? " class=\"update\"> *" : ">") + "</i></span>";
	}
	else {
		cell2.innerHTML = obj.Title != null ? obj.Title : "<i>&laquo; N/A &raquo;</i>";
	}

	if (obj.HasDescription)
		cell2.innerHTML += "<a id=\"lnkViewDescr_" + obj.ID + "\" class=\"helpIcon\" href=\"javascript: void(0)\" onclick=\"OpenProjectDescription(" + obj.ID + ");\"><img src=\"/Resources/Images/info_i.png\" alt=\"View Description\" /></a>"


	// Budget:
	var cellBd = row.insertCell(cellId++);
	cellBd.align = "right";
	cellBd.className = "nobr";
	cellBd.innerHTML = obj.BudgetString;

	// Leader/Partner:
	var cellLP = row.insertCell(cellId++);
	cellLP.align = "center";
	cellLP.innerHTML = obj.LeadPartnerStatus != null ? obj.LeadPartnerStatus : "&nbsp;";

	// Deadline (DeadlineEOIText: 5,6; DeadlineTenderText: 10,11):
	if (obj.StatusId == 5 || obj.StatusId == 6 || obj.StatusId == 10 || obj.StatusId == 11) {
		var cell6 = row.insertCell(cellId++);
		cell6.align = "right";
		cell6.innerHTML = "&nbsp;";
		if (obj.StatusId == 5 || obj.StatusId == 6)
			cell6.innerHTML = obj.DeadlineEOIText != null ? obj.DeadlineEOIText : "&nbsp;";
		else if (obj.StatusId == 10 || obj.StatusId == 11)
			cell6.innerHTML = obj.DeadlineTenderText != null ? obj.DeadlineTenderText : "&nbsp;";
		else
			cell6.style.background = "#F5F5F5";
	}

	// Entry Date (DateCreateText: 1,3):
	if (obj.StatusId == 1 || obj.StatusId == 3) {
		var cell7 = row.insertCell(cellId++);
		cell7.align = "right";
		cell7.innerHTML = "&nbsp;";
		if (obj.StatusId == 1 || obj.StatusId == 3)
			cell7.innerHTML = obj.DateCreateText != null ? obj.DateCreateText : "&nbsp;";
		else
			cell7.style.background = "#F5F5F5";
	}

	// Duration:
	if (PageTypeName == "ICATB" || PageTypeName == "ICATB-CS") {
		var cell8 = row.insertCell(cellId++);
		cell8.innerHTML = "&nbsp;";
		if (obj.Duration != null && obj.DurationMeasure != null) {
			var durationStr = obj.Duration + " ";
			if (obj.DurationMeasure == 'd')
				durationStr += "day" + (obj.Duration > 1 ? "s" : "");
			else
				durationStr += "month" + (obj.Duration > 1 ? "s" : "");
			cell8.innerHTML = durationStr;
		}
	}

	// Interested members:
	var cell4 = row.insertCell(cellId++);
	var cell4Html = "&nbsp;";
	if (obj.InterestedUsers != null) {
		if (obj.InterestedUsers.length > 0) {
			cell4Html = "<div class=\"interestedMembers compact\">";

			for (iu = 0; iu < obj.InterestedUsers.length; iu++) {
				var intrUser = obj.InterestedUsers[iu];
				var mDetailCnt = "<span id=\"imDetail_" + obj.ID + '_' + iu + "\" style=\"display:none;padding-top:5px\">" +
					(intrUser.FirstName != '' || intrUser.LastName != '' ? (intrUser.FirstName + " " + intrUser.LastName) : "<i>&laquo; Not Available &raquo;</i>") +
					"<br/>" +
					(intrUser.Email != '' ? ("<a href=\"mailto:" + intrUser.Email + "?subject=" + escape("ICANET \"" + obj.Title + "\": inquiry") + "\">" + intrUser.Email + "</a>") : '') + 
					"</span>";

				cell4Html += "<div>" +
					"<a class=\"compLink\" href=\"javascript: void(0)\" onclick=\"ShowHideElement('imDetail_" + obj.ID + '_' + iu + "')\">" +
					intrUser.CompanyName + "</a> (" + intrUser.DateTimeStr + ")" + mDetailCnt + "</div>";
			}
			cell4Html += "</div>";
		}
	}
	cell4.innerHTML = cell4Html;


	// Interest btn (render it only if project status is < 10 for ICA TB views):
	if (obj.StatusId < 10 && (PageTypeName == "ICATB" || PageTypeName == "ICATB-CS")) {
		var cell9 = row.insertCell(cellId++);
		cell9.innerHTML = "&nbsp;";
		if (obj.CanBeInterested != null && obj.CanBeInterested == true)
			cell9.innerHTML = "<nobr><a href=\"javascript: void(0)\" onclick=\"SetCompanyInterest(this.parentNode.parentNode, " + obj.ID + ")\">We too!</a></nobr>";
		else
			cell9.style.background = "#F5F5F5";
	}
	/* */
}


function SetCompanyInterest(objSender, projectId) {
	
	var svc = new ICAAjaxServices.Projects();
	svc.SaveCompanyInterestOnProject(projectId);
	// hide the button and show message:
	objSender.innerHTML = '<span style=\"color:#FF0000\">It is in <a href=\"/Intranet/MyTenderBoard\" style=\"color:#FF0000\">MY Tender Board</a> now!</span>';
}

function RemoveCompanyInterest(objSender, project2Id, projectId) {
	if (confirm("Are you sure you want to remove this project from your tenderboard?")) {
		var svc = new ICAAjaxServices.Projects();
		svc.RemoveCompanyInterestOnProject(project2Id);
		// hide the button and show text message:
		objSender.html("<span style=\"color:#FF0000\">This project will no longer exist in your Tender Board, but the data saved will remain existing in the database.</span>");
		// hide the "apply button":
		$("#btnApply").hide();
		// hide row from TB:
		parent.document.getElementById("row_" + projectId).style.display = "none";
	}
}


function AddConflictOfInterest(objSender, projectId, resultContainer) {
	var svc = new ICAAjaxServices.Projects();
	
	var comment = "";

	svc.AddConflictOfInterest(projectId, comment, AddConflictOfInterest_OnSuccess, null, resultContainer);

	// show message:
	$("#" + objSender.id + "_alert").show();
}


function AddConflictOfInterest_OnSuccess(result, userctx) {
	
	if (userctx != null && userctx != '') {
		$("#" + userctx).val(result);
	}
}


function CancelConflictOfInterest(objSender, conflictId, projectId, resultContainer) {
	var svc = new ICAAjaxServices.Projects();

	var comment = "";
	
	svc.CancelConflictOfInterest(conflictId, projectId, comment, CancelConflictOfInterest_OnSuccess, null, resultContainer);

	// hide message:
	$("#" + objSender.id + "_alert").hide();
}

function CancelConflictOfInterest_OnSuccess(result, userctx) {

	// empty the conflict ID value:
	if (userctx != null && userctx != '')
		$("#" + userctx).val(0);
}


function LoadDocumentList(project2Id, containerId) {
	var svc = new ICAAjaxServices.Projects();
	
	svc.RetrieveDocumentList(CurrentProject2Id, "divDocumentsList", LoadDocumentList_OnSuccess, null, null);
}

function LoadDocumentList_OnSuccess(result) {
	
	if (result.DocumentList == null || result.DocumentList.length == 0) {
		$('#' + result.ContainerId).html("<i>Currently there are no documents uploaded.</i>");
		return;
	}
	$('#' + result.ContainerId).html("");
	for (i = 0; i < result.DocumentList.length; i++) {
		DocumentList_AddItemRow(result.ContainerId, result.DocumentList[i]);
	}
}

function DocumentList_AddItemRow(containerId, obj) {
	document.getElementById(containerId).innerHTML += "<a href=\"/GetDocument/?id=" + obj.ID + "\" target=\"_blank\">" + obj.Title + "</a> (" + obj.TypeName + ")<br />";
}

function OpenProjectDescription(projectId) {
	var svc2 = new ICAAjaxServices.Projects();
	svc2.GetProjectDescription(projectId, OpenProjectDescription_OnSuccess, null, null);
}

function OpenProjectDescription_OnSuccess(result) {
	if (result != '') {
		$("#divPopupContent").html(result);
		ShowPopup('/info.htm?src=divPopupContent', 'projectdescr', 600, 500, true, true);
	}
	else
		alert("Description not available");
}

function SetupTBList(hdrName, bShow) {
	if (bShow) {
		$('#hdr' + hdrName).show();
		$('#count' + hdrName).show();
		$('#tbl' + hdrName).show();
		// clean table when setting to "visible":
		$("#tbl" + hdrName + " tr:gt(0)").remove();
	}
	else {
		$('#hdr' + hdrName).hide();
		$('#count' + hdrName).hide();
		$('#tbl' + hdrName).hide();
	}
}


var StatusIdsTB = new Array();
StatusIdsTB["tblProjectsList"] = null;
StatusIdsTB["tblIdentificList"] = 1;
StatusIdsTB["tblPipeActionList"] = 3;
StatusIdsTB["tblEOIPreparationList"] = 5;
StatusIdsTB["tblEOIAwaitingList"] = 6;
StatusIdsTB["tblTenderPreparationList"] = 10;
StatusIdsTB["tblTenderAwaitingList"] = 11;
StatusIdsTB["tblProjectsClosedList"] = 31;
StatusIdsTB["tblProjectsWonList"] = 20;
StatusIdsTB["tblProjectPreparationList"] = 21;
StatusIdsTB["tblInceptionList"] = 22;
StatusIdsTB["tblProjRunningList"] = 23;
StatusIdsTB["tblProjClosingList"] = 24;
StatusIdsTB["tblContrClosureList"] = 30;
StatusIdsTB["tblFinalDebriefList"] = 40;
StatusIdsTB["tblIdentifNotFollowedList"] = 2;
StatusIdsTB["tblPipeNotFollowedList"] = 4;
StatusIdsTB["tblEoiCancelledList"] = 7;
StatusIdsTB["tblEoiNotFollowedList"] = 8;
StatusIdsTB["tblEoiLostList"] = 9;
StatusIdsTB["tblTenderCancelledList"] = 17;
StatusIdsTB["tblTenderNotFollowedList"] = 18;
StatusIdsTB["tblTenderLostList"] = 19;
StatusIdsTB["tblContrCancelledList"] = 26;
StatusIdsTB["tblProjSuspendedList"] = 27;
StatusIdsTB["tblProjCancelledList"] = 28;
StatusIdsTB["tblProjDroppedList"] = 29;