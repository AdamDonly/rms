var selectedSectorIDs = "";

function ScrollToMainSector(mainSectorId, sectorsDivId) {
	var mainSector = $("#mainSector_" + mainSectorId)[0];
	var sectorDiv = $("#" + sectorsDivId)[0];
	sectorDiv.scrollTop = mainSector.offsetTop;
}

function CheckMainSector(id, storageId, displayId) {
	var mainSectorObj = $('#mainSector_' + id)[0];
	var sectors = $('#subSectors_' + id + ' input[type="checkbox"]');
	var selectedIdsList = PARENT.document.getElementById(storageId).value;
	for (i = 0; i < sectors.length; i++) {
		// addit only if not selected:
		sectors[i].checked = mainSectorObj.checked && selectedIdsList.indexOf(sectors[i].value) == -1;

		CheckSector(sectors[i].value, storageId);

	}
}

function CheckSector(id, storageId) {

	var sectorObj = $('#sector_' + id)[0];
	if (sectorObj.checked && selectedSectorIDs.indexOf(id) == -1) {
		if (selectedSectorIDs != "") selectedSectorIDs += ",";
		selectedSectorIDs += id;
	}
	else if (!sectorObj.checked && selectedSectorIDs.indexOf(id) > -1) {
		var removeSectorStr = id;
		// check if the id is first in the list and there are more behind it:
		if (selectedSectorIDs.indexOf(id + ",") > -1) removeSectorStr = removeSectorStr + ",";
		// check if there's a comma in front of the id:
		else if (selectedSectorIDs.indexOf("," + id) > -1) removeSectorStr = "," + removeSectorStr;

		selectedSectorIDs = selectedSectorIDs.replace(removeSectorStr, "");
	}
}

function SynchronizeDisplaySectorList(storageId, displayId) {
	var selectedIdsObj = PARENT.document.getElementById(storageId);
	var selectedIdsDisplayObj = PARENT.document.getElementById(displayId);
	selectedIdsObj.value = selectedSectorIDs;

	var selectedIdsSplit = selectedSectorIDs.split(',');
	if (selectedSectorIDs != "" && selectedIdsSplit.length > 0) {
		selectedIdsDisplayObj.innerHTML = "";
		for (m = 0; m < selectedIdsSplit.length; m++) {
			if (selectedIdsSplit[m] != "") {
				PARENT.BuildHtmlListItem(displayId, selectedIdsSplit[m], $("#sector_" + selectedIdsSplit[m])[0].title, storageId);
			//	selectedIdsDisplayObj.innerHTML +=
			//		(selectedIdsDisplayObj.innerHTML != "" ? "<br/>" : "") +
			//		$("#sector_" + selectedIdsSplit[m])[0].title;
			}
		}
	}
	else {
		selectedIdsDisplayObj.innerHTML = "--- No selection ---";
	}
}

function ApplySectors(storageId, displayId, isInNewWnd) {
	SynchronizeDisplaySectorList(storageId, displayId);

	if (isInNewWnd)
		self.close();
	else
		PARENT.HideHtmlPopup();
}


function SetCheckedSectors(parentElementId) {
	// parent
	//		|___
	//			mainsector (id = 'mainSector_<MainSectorId>')
	//		|___
	//			subsectors DIV (id = 'subSectors_<MainSectorId>')
	//												|___
	//													sector (id = 'sector_<sectorId>')

	var mainSectors = $("#" + parentElementId + " input[type='checkbox']");
	for (i = 0; i < mainSectors.length; i++) {
		if (mainSectors[i].id.indexOf("mainSector_") == -1) continue;

		var mainSectorDbId = mainSectors[i].id.replace("mainSector_", "");
		var subSectors = $("#subSectors_" + mainSectorDbId + "  input[type='checkbox']");
		var cntSelectedSubSectors = 0;
		for (k = 0; k < subSectors.length; k++) {
			var subSectorDbId = subSectors[k].id.replace("sector_", "");
			if (selectedSectorIDs.indexOf(subSectorDbId) > -1)
			{
				// should be checked:
				subSectors[k].checked = true;
				$("#" + mainSectors[i].id + "_leftTitle").css("font-weight", "bold");
			//	$("#" + mainSectors[i].id + "_rightTitle").css("font-weight", "bold");
				cntSelectedSubSectors++;
			}
		}

		if (cntSelectedSubSectors > 0) {
			if (cntSelectedSubSectors == subSectors.length) {
				// check also the main sector:
				mainSectors[i].checked = true;
			}

			$("#" + mainSectors[i].id + "_leftTitle").html($("#" + mainSectors[i].id + "_leftTitle").html() + " (" + cntSelectedSubSectors + ")");
		//	$("#" + mainSectors[i].id + "_rightTitle").html($("#" + mainSectors[i].id + "_rightTitle").html() + " (" + cntSelectedSubSectors + ")");
		}
	}
}