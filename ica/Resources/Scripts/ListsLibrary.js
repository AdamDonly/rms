
var itemPageArray = new Array();
var itemFinishedArray = new Array();

var ListCurrentPageNumber = 1;
var LockListExecution = 0;

function GetPageLinkHtml(tableId, linkText, goToPageNum, cssClass) {
	return "<a " + (cssClass != '' ? "class=\"" + cssClass + "\"" : "") + 
		"onclick=\"List_GoToPage('" + tableId + "', " + goToPageNum + ")\" href=\"javascript:void(0)\">" + linkText + "</a>";
}

function List_AddPagingRow(tableId, currentPage, totalPages) {

	var pgContainer = document.getElementById(tableId.replace("tbl", "pages"));
	if (pgContainer != null) pgContainer.innerHTML = "";

	if (totalPages < 2) return;

	if (currentPage > 1) {
		pgContainer.innerHTML =
			GetPageLinkHtml(tableId, "first", 1, "first") +
			GetPageLinkHtml(tableId, "previous", currentPage - 1, "prev");
	}

	// render a total of 7 (most close) pages:
	var startPg = (currentPage - Math.floor((7 / 2))) < 1 ? 1 : (currentPage - Math.floor((7 / 2)));
	var endPg = startPg + 7 - 1;
	if (endPg > totalPages) endPg = totalPages;
	for (i = startPg; i <= endPg; i++) {
		pgContainer.innerHTML +=
			GetPageLinkHtml(tableId, i, i, (i == parseInt(currentPage) ? "currentPage" : ""));
	}

	if (currentPage < totalPages) {
		pgContainer.innerHTML +=
			GetPageLinkHtml(tableId, "next", currentPage + 1, "next") +
			GetPageLinkHtml(tableId, "last", totalPages, "last");
	}
}

function ListTB_AddPagingRow(tableId, currentPage, totalPages, projectStatus) {
	urlParams.Status = StatusIdsTB[tableId][0];
	List_AddPagingRow(tableId, currentPage, totalPages);
}

// only page number and arrows for fwd and back:
function List_AddShortPagingRow(tableId, currentPage, totalPages) {
	if (totalPages < 2) return;

	var table = document.getElementById(tableId);

	var row = table.insertRow(table.rows.length);

	var cell = row.insertCell(0);
	cell.className = "pagingCell";
	cell.colSpan = ListTableColumnsCount;

	if (currentPage > 1) {
		// render the "back" link:
		var link = document.createElement('a');
		link.href = "javascript: void(0)";
		link.setAttribute("onclick", "List_GoToPage('" + tableId + "', " + (currentPage - 1) + ")");
		link.innerHTML = "&lt;&lt;";
		cell.appendChild(link);
	}


	// render current page number:
	var currentPageMark = document.createElement('a');
	currentPageMark.innerHTML = currentPage;
	currentPageMark.className = "currentPage";
	cell.appendChild(currentPageMark);

	if (currentPage < totalPages) {
		// render the "forward" link:
		var link = document.createElement('a');
		link.href = "javascript: void(0)";
		link.setAttribute("onclick", "List_GoToPage('" + tableId + "', " + (currentPage + 1) + ")");
		link.innerHTML = "&gt;&gt;";
		cell.appendChild(link);
	}
}


function List_AddNoResultsRow(tableId) {
	var table = document.getElementById(tableId);

	var row = table.insertRow(table.rows.length);

	var cell = row.insertCell(0);
	cell.className = "nodata";
	cell.colSpan = ListTableColumnsCount;
	cell.innerHTML = "No results found.";
}


function List_GoToPage(tblId, pageNumber) {
	ListCurrentPageNumber = pageNumber;

	List_Update(tblId, false);
}

function List_ResetPageCounter() {
	ListCurrentPageNumber = 1;
}


function List_AppendPage(itemId, pageNumber) {
	ListCurrentPageNumber = pageNumber;

	List_Append("tbl" + itemId);
}