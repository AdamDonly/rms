var selectedCountryIDs = "";

function ScrollToRegion(regionId, countriesDivId) {
	var region = $("#region_" + regionId)[0];
	var countriesDiv = $("#" + countriesDivId)[0];
	countriesDiv.scrollTop = region.offsetTop;
}

function CheckRegion(id, storageId, displayId) {
	var regionObj = $('#region_' + id)[0];
	var countries = $('#countries_' + id + ' input[type="checkbox"]');
	var selectedIdsList = PARENT.document.getElementById(storageId).value;
	for (i = 0; i < countries.length; i++) {
		// addit only if not selected:
		countries[i].checked = regionObj.checked && selectedIdsList.indexOf(countries[i].value) == -1;

		CheckCountry(countries[i].value, storageId);

	}
}

function CheckCountry(id, storageId) {

	var countryObj = $('#country_' + id)[0];
	if (countryObj.checked && selectedCountryIDs.indexOf(id) == -1) {
		if (selectedCountryIDs != "") selectedCountryIDs += ",";
		selectedCountryIDs += id;
	}
	else if (!countryObj.checked && selectedCountryIDs.indexOf(id) > -1) {
		var removeCountryStr = id;
		// check if the id is first in the list and there are more behind it:
		if (selectedCountryIDs.indexOf(id + ",") > -1) removeCountryStr = removeCountryStr + ",";
		// check if there's a comma in front of the id:
		else if (selectedCountryIDs.indexOf("," + id) > -1) removeCountryStr = "," + removeCountryStr;

		selectedCountryIDs = selectedCountryIDs.replace(removeCountryStr, "");
	}
}

function SynchronizeDisplayCountryList(storageId, displayId) {
	var selectedIdsObj = PARENT.document.getElementById(storageId);
	var selectedIdsDisplayObj = PARENT.document.getElementById(displayId);
	selectedIdsObj.value = selectedCountryIDs;

	var selectedIdsSplit = selectedCountryIDs.split(',');
	if (selectedCountryIDs != "" && selectedIdsSplit.length > 0) {
		selectedIdsDisplayObj.innerHTML = "";
		for (m = 0; m < selectedIdsSplit.length; m++) {
			if (selectedIdsSplit[m] != "") {
				PARENT.BuildHtmlListItem(displayId, selectedIdsSplit[m], $("#country_" + selectedIdsSplit[m])[0].title, storageId);
			//	selectedIdsDisplayObj.innerHTML +=
			//		(selectedIdsDisplayObj.innerHTML != "" ? "<br/>" : "") +
			//		$("#country_" + selectedIdsSplit[m])[0].title;
			}
		}
	}
	else {
		selectedIdsDisplayObj.innerHTML = "--- No selection ---";
	}
}

function ApplyCountries(storageId, displayId, isInNewWnd) {
	SynchronizeDisplayCountryList(storageId, displayId);

	if (isInNewWnd)
		self.close();
	else
		PARENT.HideHtmlPopup();
}


function SetCheckedCountries(parentElementId) {

	var regions = $("#" + parentElementId + " input[type='checkbox']");
	for (i = 0; i < regions.length; i++) {
		if (regions[i].id.indexOf("region_") == -1) continue;

		var regionDbId = regions[i].id.replace("region_", "");
		var countries = $("#countries_" + regionDbId + "  input[type='checkbox']");
		var cntSelectedCountries = 0;
		for (k = 0; k < countries.length; k++) {
			var countryDbId = countries[k].id.replace("country_", "");
			if (selectedCountryIDs.indexOf(countryDbId) > -1)
			{
				// should be checked:
				countries[k].checked = true;
				$("#" + regions[i].id + "_leftTitle").css("font-weight", "bold");
				cntSelectedCountries++;
			}
		}

		if (cntSelectedCountries > 0) {
			if (cntSelectedCountries == countries.length) {
				// check also the region:
				regions[i].checked = true;
			}

			$("#" + regions[i].id + "_leftTitle").html($("#" + regions[i].id + "_leftTitle").html() + " (" + cntSelectedCountries + ")");
			//	$("#" + mainSectors[i].id + "_rightTitle").html($("#" + mainSectors[i].id + "_rightTitle").html() + " (" + cntSelectedSubSectors + ")");
		}
	}
}