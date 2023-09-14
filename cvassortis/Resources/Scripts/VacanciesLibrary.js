
function LoadPositionList(
	projectType, sectorList, region, longTerm, shortTerm, pageSize, pageNumber, containerId) {
	var svc = new ICAAjaxServices.Vacancies();

	svc.GetPositionsList(projectType, sectorList, region, longTerm, shortTerm, pageSize, pageNumber, containerId, LoadPositionList_OnSuccess, null, null);
}

function LoadPositionList_OnSuccess(result) {

	$("#vacResultsCount").html("");
	if (result.PositionList == null || result.PositionList.length == 0) {
		$("#vacResultsCount").html("0 vacancies");
		List_AddNoResultsRow(result.ContainerId);
		return;
	}

	//	List_AddShortPagingRow(result.ContainerId, result.CurrentPage, Math.ceil(result.PositionsTotalCount / result.PageSize));

	$("#vacResultsCount").html(result.TotalCount + " vacanc" + (result.TotalCount == 1 ? "y" : "ies"));

	for (i = 0; i < result.PositionList.length; i++) {
		PositionList_AddItemRow(result.ContainerId, result.PositionList[i]);
	}

	List_AddPagingRow(result.ContainerId, result.CurrentPage, Math.ceil(result.TotalCount / result.PageSize));

//	document.location = "#tableTop";
}

function List_Update(tblId, resetPageCounter) {
	// clear table:
	$("#" + tblId + " tr:gt(0)").remove();

	if (resetPageCounter)
		List_ResetPageCounter();

	var projectType = "";
	var sectorAbbr = $("#vacancy_sectors") != null ? $("#vacancy_sectors").children("option:selected")[0].value : "";
	var region = $("#vacancy_countries") != null ? $("#vacancy_countries").children("option:selected")[0].value : "";
	var longTerm = null;
	var shortTerm = null;

	LoadPositionList(projectType, sectorAbbr, region, longTerm, shortTerm, ListPageSize, ListCurrentPageNumber, "tblVacanciesList");
}


function PositionList_AddItemRow(tableId, obj) {

	var table = document.getElementById(tableId);

	var row = table.insertRow(table.rows.length);
	row.onmouseover = function () { this.style.backgroundColor = '#f3f3f3'; };
	row.onmouseout = function () { this.style.backgroundColor = '#fff'; };

	// Title:
	var cell0 = row.insertCell(0);
	cell0.innerHTML = "<a href=\"" + 
		(obj.SourceDB == 'ibf' ? IBF_VacancyDetailURL : ICA_VacancyDetailURL) + obj.ID + "\">" + obj.Title + "</a>";

	// Country:
	var cell1 = row.insertCell(1);
	cell1.innerHTML = obj.Country != null ? obj.Country : "&nbsp;";

	// Project Title:
	var cell2 = row.insertCell(2);
	cell2.innerHTML = obj.Project != null ? obj.Project.Title : "&nbsp;";

	// Workdays:
	var cell3 = row.insertCell(3);
	cell3.innerHTML = obj.Workdays != null ? obj.Workdays : "&nbsp;";

	// Estimated to start:
	var cell4 = row.insertCell(4);
	cell4.innerHTML = obj.AvailabilityString != null ? obj.AvailabilityString : "&nbsp;";

	// Application Deadline:
	var cell5 = row.insertCell(5);
	cell5.innerHTML = obj.DeadlineStr;

	// Company:
	var cell6 = row.insertCell(6);
	cell6.innerHTML = obj.CompanyName != null ? obj.CompanyName : "&nbsp;";
}
