
function LoadReferenceList(pageSize, pageNumber, containerId) {

	var svc = new ICAAjaxServices.References();

	svc.GetReferencesList(
		urlParams.Keyword,
		urlParams.Donor,
		urlParams.Country,
		urlParams.SectorList,
		urlParams.MainSector,
		urlParams.Unit,
		urlParams.BudgetMin,
		urlParams.BudgetMax,
		urlParams.YearFrom,
		urlParams.YearTo,
		urlParams.ProportionWdMin,
		urlParams.ProportionWdMax,
		urlParams.ProportionBaMin,
		urlParams.ProportionBaMax,
		urlParams.ProportionCsMin,
		urlParams.ProportionCsMax,
		urlParams.IncludeNonAwarded,
		urlParams.SearchAllCompanies,
		pageSize,
		pageNumber,
		containerId, 
		LoadReferenceList_OnSuccess, null, null);
}

function LoadReferenceList_OnSuccess(result) {

	// release lock:
	LockListExecution = 0;

	// update location hash (for "back" button compatibility):
	var paramsLocation = urlParams.assembleUrlHash();
	$("paramsHolder").attr("name", paramsLocation.replace("#", ""));
	document.location = paramsLocation;

	// mark the list as "finished", to not be appended anymore:
	if (result.ReferenceList.length < result.PageSize) {
		itemFinishedArray[result.ContainerId] = 1;
	}

	if ((result.ReferenceList == null || result.ReferenceList.length == 0) && $("#" + result.ContainerId + " tr").size() < 3) {
		List_AddNoResultsRow(result.ContainerId);
		return;
	}

	for (i = 0; i < result.ReferenceList.length; i++) {
		ReferenceList_AddItemRow(result.ContainerId, result.ReferenceList[i]);
	}

//	List_AddPagingRow(result.ContainerId, result.CurrentPage, Math.ceil(result.ReferencesTotalCount / result.PageSize));
}

function List_Update(tblId, resetPageCounter) {

	// delay execution if locked:
	if (LockListExecution != 0) {
		setTimeout("List_Update('" + tblId + "', " + resetPageCounter + ")", 500);
		return;
	}

	// lock execution:
	LockListExecution = 1;

	// clear table:
	$("#" + tblId + " tr:gt(1)").remove();

	if (resetPageCounter)
		List_ResetPageCounter();

	LoadReferenceList(ListPageSize, ListCurrentPageNumber, tblId);
}

function List_Append(tblId) {
	// delay execution if locked:
	if (LockListExecution != 0) {
		setTimeout("List_Append('" + tblId + "')", 500);
		return;
	}

	// lock execution:
	LockListExecution = 1;

	// include 'not awarded', because there are such statuses:
	if (urlParams.Status >= 1 && urlParams.Status < 20
		|| urlParams.Status > 24 && urlParams.Status != 30 && urlParams.Status != 31 && urlParams.Status != 40) {
		urlParams.IncludeNonAwarded = true;
	}

	LoadReferenceList(ListPageSize, ListCurrentPageNumber, tblId);
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

	urlParams.Keyword = $("#keywords").attr("value");
	urlParams.Donor = $("#reference_fundingAgencies") != null && $("#reference_fundingAgencies").children("option:selected").length > 0 ? $("#reference_fundingAgencies").children("option:selected")[0].value : "";
	urlParams.Country = $("#project_countries") != null && $("#project_countries").children("option:selected").length > 0 ? $("#project_countries").children("option:selected")[0].value : "";
	urlParams.SectorList = $("#hidSectors").attr("value") != "" ? $("#hidSectors").attr("value") : null;
	urlParams.Unit = "";
	urlParams.BudgetMin = parseInt($("#budget_from").attr("value")) > 0 ? parseInt($("#budget_from").attr("value")) : null;
	urlParams.BudgetMax = parseInt($("#budget_to").attr("value")) > 0 ? parseInt($("#budget_to").attr("value")) : null;
	urlParams.YearFrom = parseInt($("#period_from").attr("value")) > 0 ? parseInt($("#period_from").attr("value")) : null;
	urlParams.YearTo = parseInt($("#period_to").attr("value")) > 0 ? parseInt($("#period_to").attr("value")) : null;

	urlParams.ProportionWd = $("#chkProportionWD").attr("checked") == true;
	urlParams.ProportionBa = $("#chkProportionBA").attr("checked") == true;
	urlParams.ProportionCs = $("#chkProportionCS").attr("checked") == true;

	urlParams.ProportionWdMin = urlParams.ProportionWd && parseInt($("#proportion_from").attr("value")) > 0 ? parseInt($("#proportion_from").attr("value")) : null;
	urlParams.ProportionWdMax = null;
	urlParams.ProportionBaMin = urlParams.ProportionBa && parseInt($("#proportion_from").attr("value")) > 0 ? parseInt($("#proportion_from").attr("value")) : null;
	urlParams.ProportionBaMax = null;
	urlParams.ProportionCsMin = urlParams.ProportionCs && parseInt($("#proportion_from").attr("value")) > 0 ? parseInt($("#proportion_from").attr("value")) : null;
	urlParams.ProportionCsMax = null;

//	urlParams.IncludeNonAwarded = document.getElementById("cbIncludeNotAwarded").checked;
	urlParams.SearchAllCompanies = document.getElementById("cbSearchAllCompanies").checked;
}


function ReferenceList_AddItemRow(tableId, obj) {

	var referenceList = $("#selref")[0].value; //TODO: fill it somehow when it is clear for what it is used :)

	var table = document.getElementById(tableId);

	var row = table.insertRow(table.rows.length);
	row.onmouseover = function () { this.style.backgroundColor = '#f3f3f3'; };
	row.onmouseout = function () { this.style.backgroundColor = '#fff'; };

	// Company name:
	var cell0 = row.insertCell(0);
	cell0.innerHTML = obj.ICAMemberName != null ? obj.ICAMemberName : "&nbsp;";

	// Country:
	var cell1 = row.insertCell(1);
	cell1.innerHTML = obj.Country != null ? obj.Country : "&nbsp;";

	// Name:
	var cell2 = row.insertCell(2);
	cell2.innerHTML = '<a href="CompileFiche/?selfmt=ec&selref=' + obj.DefaultLanguage.toLowerCase() + obj.Project2Id + '" target="_blank">' + obj.Title + '</a>';

	// Funding:
	var cell3 = row.insertCell(3);
	cell3.innerHTML = obj.Funding != null ? obj.Funding : "&nbsp;";

	// Period:
	var cell4 = row.insertCell(4);
	if (obj.StartDate != null && obj.EndDate != null) {
		var startDate_Month = obj.StartDate.getMonth() + 1;
		if (startDate_Month < 10) startDate_Month = '0' + startDate_Month;

		var startDate_year = obj.StartDate.getYear() + '';
		if (startDate_year.length == 4)
			startDate_year = startDate_year.substr(2, 2);
		
		var endDate_Month = obj.EndDate.getMonth() + 1;
		if (endDate_Month < 10) endDate_Month = '0' + endDate_Month;

		var endDate_year = obj.EndDate.getYear() + '';
		if (endDate_year.length == 4)
			endDate_year = endDate_year.substr(2, 2);

		cell4.innerHTML = '<nobr>' + startDate_Month + '/' + startDate_year + ' - ' + endDate_Month + '/' + endDate_year + '</nobr>';
	}
	else
		cell4.innerHTML = "&nbsp;";

	// Budget:
	var cell5 = row.insertCell(5);
	cell5.align = "right";
	if (obj.Budget != null && obj.Budget != '')
		cell5.innerHTML = obj.Budget;
	else
		cell5.innerHTML = obj.BudgetAmount != null && obj.Currency != null ? "<nobr>" + FormatNumber(obj.BudgetAmount, ',', 2, '.') + " " + obj.Currency + "</nobr>" : "&nbsp;";

	// Proportion Working Days:
	var cell6 = row.insertCell(6);
	cell6.align = "right";
	cell6.innerHTML = obj.ProportionDays != null ? "<nobr>" + obj.ProportionDays + "%</nobr>" : "&nbsp;";

	// Proportion Budget Amount:
	var cell7 = row.insertCell(7);
	cell7.align = "right";
	cell7.innerHTML = obj.ProportionAmount != null ? "<nobr>" + obj.ProportionAmount + "%</nobr>" : "&nbsp;";

	// Proportion Contract share:
	var cell8 = row.insertCell(8);
	cell8.align = "right";
	cell8.innerHTML = obj.ProportionAgreed != null ? "<nobr>" + obj.ProportionAgreed + "%</nobr>" : "&nbsp;";

	// Checkbox EN:
	var cell9 = row.insertCell(9);
	cell9.align = "center";
	cell9.className = "printHidden";
	cell9.innerHTML = obj.English == 1
	? ("<input type=\"checkbox\" id=\"en" + obj.Project2Id + "\" name=\"en" + obj.Project2Id + "\" " +
		(referenceList.indexOf(("_en" + obj.Project2Id)) > -1 ? " checked " : "") +
		"onclick=\"javascript: CheckReference(" + obj.Project2Id + ", 'en');\" />")
	: "&nbsp;";

	// Checkbox FR:
	var cell10 = row.insertCell(10);
	cell10.align = "center";
	cell10.className = "printHidden";
	cell10.innerHTML = obj.French == 1
	? ("<input type=\"checkbox\" id=\"fr" + obj.Project2Id + "\" name=\"fr" + obj.Project2Id + "\" " +
		(referenceList.indexOf(("_fr" + obj.Project2Id)) > -1 ? " checked " : "") +
		"onClick=\"CheckReference(" + obj.Project2Id + ", 'fr');\" />")
	: "&nbsp;";

	// Checkbox ES:
	var cell11 = row.insertCell(11);
	cell11.align = "center";
	cell11.className = "printHidden";
	cell11.innerHTML = obj.Spanish == 1
	? ("<input type=\"checkbox\" id=\"es" + obj.Project2Id + "\" name=\"es" + obj.Project2Id + "\" " +
		(referenceList.indexOf(("_es" + obj.Project2Id)) > -1 ? " checked " : "") +
		"onClick=\"CheckReference(" + obj.Project2Id + ", 'es');\" />")
	: "&nbsp;";

}


function CheckReference(num, lng) {
	var pos;
	var eids;
	var enumstr;
	enumstr = '_' + lng + '' + num;
	eids = $("#selref").attr("value");

	if ($("#" + lng + num)[0].checked == true) {
		$("#selref")[0].value = eids + enumstr;
	}
	else {
		pos = eids.indexOf(enumstr);
		if (pos > -1)
			$("#selref")[0].value = eids.substring(0, pos) + eids.substring(pos + enumstr.length, eids.length);
	}

//	document.cookie = 'Reference=' + $("#selref").value;
}

function CheckFormat(objSrc) {
	$("#selfmt")[0].value = objSrc.options[objSrc.selectedIndex].value;
//	document.cookie = 'ReferenceFormat=' + objSrc.options[objSrc.selectedIndex].value;
}

function DeleteReference(refId) {
	var svc = new ICAAjaxServices.References();
	svc.DeleteReference(refId, DeleteReference_OnSuccess, null, null);
}

function DeleteReference_OnSuccess(result) {
	if (result) {
		// reload the references list:
		eval(DeleteReference_OnSuccessAction);
	}
	else
		alert("not deleted");
}