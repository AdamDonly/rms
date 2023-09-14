
function URLParams() {

	this.Keys = new URLParamKeys();

	this.CurrentPage = 1;
	this.ProjectType = null;
	this.Keyword = "";
	this.Donor = "";
	this.DonorId = null;
	this.Country = "";
	this.Region = "";
	this.SectorList = null;
	this.MainSector = null;
	this.Status = null;
	this.Unit = "";
	this.BudgetMin = null;
	this.BudgetMax = null;
	this.Orderby = "";
	this.YearFrom = null;
	this.YearTo = null;
	this.ProportionWd = true;
	this.ProportionWdMin = null;
	this.ProportionWdMax = null;
	this.ProportionBa = true;
	this.ProportionBaMin = null;
	this.ProportionBaMax = null;
	this.ProportionCs = true;
	this.ProportionCsMin = null;
	this.ProportionCsMax = null;
	this.LongTerm = null;
	this.ShortTerm = null;
	this.CompanyId = null;
	this.GetCompanyProjectsOnly = null;
	this.IncludeNonAwarded = null;
	this.SearchAllCompanies = null;

	// init the variables from URL:
	var paramPairs = document.location.hash.toLowerCase().replace("#", "").split(";");

	for (s = 0; s < paramPairs.length; s++) {

		if (paramPairs[s] == "") continue;

		var paramPairSpl = paramPairs[s].split(":");

		if (paramPairSpl.length < 2) continue;

		switch (paramPairSpl[0]) {
			case this.Keys.CurrentPage:
				this.CurrentPage = paramPairSpl[1];
				break;

			case this.Keys.ProjectType:
				this.ProjectType = paramPairSpl[1];
				break;

			case this.Keys.Keyword:
				this.Keyword = paramPairSpl[1];
				break;

			case this.Keys.Donor:
				this.Donor = paramPairSpl[1];
				break;

			case this.Keys.DonorId:
				this.DonorId = paramPairSpl[1];
				break;

			case this.Keys.Country:
				this.Country = paramPairSpl[1];
				break;

			case this.Keys.Region:
				this.Region = paramPairSpl[1];
				break;

			case this.Keys.SectorList:
				this.SectorList = paramPairSpl[1];
				break;

			case this.Keys.MainSector:
				this.MainSector = paramPairSpl[1];
				break;

			case this.Keys.Status:
				this.Status = paramPairSpl[1];
				break;

			case this.Keys.Unit:
				this.Unit = paramPairSpl[1];
				break;

			case this.Keys.BudgetMin:
				this.BudgetMin = paramPairSpl[1];
				break;

			case this.Keys.BudgetMax:
				this.BudgetMax = paramPairSpl[1];
				break;

			case this.Keys.Orderby:
				this.Orderby = paramPairSpl[1];
				break;

			case this.Keys.YearFrom:
				this.YearFrom = paramPairSpl[1];
				break;

			case this.Keys.YearTo:
				this.YearTo = paramPairSpl[1];
				break;

			case this.Keys.ProportionWd:
				this.ProportionWd = paramPairSpl[1];
				break;

			case this.Keys.ProportionWdMin:
				this.ProportionWdMin = paramPairSpl[1];
				break;

			case this.Keys.ProportionWdMax:
				this.ProportionWdMax = paramPairSpl[1];
				break;

			case this.Keys.ProportionBa:
				this.ProportionBa = paramPairSpl[1];
				break;

			case this.Keys.ProportionBaMin:
				this.ProportionBaMin = paramPairSpl[1];
				break;

			case this.Keys.ProportionBaMax:
				this.ProportionBaMax = paramPairSpl[1];
				break;

			case this.Keys.ProportionCs:
				this.ProportionCs = paramPairSpl[1];
				break;

			case this.Keys.ProportionCsMin:
				this.ProportionCsMin = paramPairSpl[1];
				break;

			case this.Keys.ProportionCsMax:
				this.ProportionCsMax = paramPairSpl[1];
				break;

			case this.Keys.LongTerm:
				this.LongTerm = paramPairSpl[1];
				break;

			case this.Keys.ShortTerm:
				this.ShortTerm = paramPairSpl[1];
				break;

			case this.Keys.CompanyId:
				this.CompanyId = paramPairSpl[1];
				break;

			case this.Keys.GetCompanyProjectsOnly:
				this.GetCompanyProjectsOnly = paramPairSpl[1];
				break;

			case this.Keys.IncludeNonAwarded:
				this.IncludeNonAwarded = paramPairSpl[1];
				break;

			case this.Keys.SearchAllCompanies:
				this.SearchAllCompanies = paramPairSpl[1];
				break;
		}
	}

	this.assembleUrlHash = function () {
		
		// params seprarator - ";"
		// name and value separator - ":"
		// multiple values separator - ","
		// exapmple: #p:1;sect:10,12,15;a:1
		var retHash = "";
		if (!IsNullOrEmpty(this.CurrentPage))
			retHash += (retHash != "" ? ";" : "") + this.Keys.CurrentPage + ":" + this.CurrentPage;

		if (!IsNullOrEmpty(this.ProjectType))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProjectType + ":" + this.ProjectType;

		if (!IsNullOrEmpty(this.Keyword))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Keyword + ":" + this.Keyword;

		if (!IsNullOrEmpty(this.Donor))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Donor + ":" + this.Donor;

		if (!IsNullOrEmpty(this.DonorId))
			retHash += (retHash != "" ? ";" : "") + this.Keys.DonorId + ":" + this.DonorId;

		if (!IsNullOrEmpty(this.Country))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Country + ":" + this.Country;

		if (!IsNullOrEmpty(this.Region))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Region + ":" + this.Region;

		if (!IsNullOrEmpty(this.SectorList))
			retHash += (retHash != "" ? ";" : "") + this.Keys.SectorList + ":" + this.SectorList;

		if (!IsNullOrEmpty(this.MainSector))
			retHash += (retHash != "" ? ";" : "") + this.Keys.MainSector + ":" + this.MainSector;

		if (!IsNullOrEmpty(this.Status))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Status + ":" + this.Status;

		if (!IsNullOrEmpty(this.Unit))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Unit + ":" + this.Unit;

		if (!IsNullOrEmpty(this.BudgetMin))
			retHash += (retHash != "" ? ";" : "") + this.Keys.BudgetMin + ":" + this.BudgetMin;

		if (!IsNullOrEmpty(this.BudgetMax))
			retHash += (retHash != "" ? ";" : "") + this.Keys.BudgetMax + ":" + this.BudgetMax;

		if (!IsNullOrEmpty(this.Orderby))
			retHash += (retHash != "" ? ";" : "") + this.Keys.Orderby + ":" + this.Orderby;

		if (!IsNullOrEmpty(this.YearFrom))
			retHash += (retHash != "" ? ";" : "") + this.Keys.YearFrom + ":" + this.YearFrom;

		if (!IsNullOrEmpty(this.YearTo))
			retHash += (retHash != "" ? ";" : "") + this.Keys.YearTo + ":" + this.YearTo;

		if (!this.ProportionWd)
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionWd + ":" + this.ProportionWd;

		if (!IsNullOrEmpty(this.ProportionWdMin))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionWdMin + ":" + this.ProportionWdMin;

		if (!IsNullOrEmpty(this.ProportionWdMax))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionWdMax + ":" + this.ProportionWdMax;

		if (!this.ProportionBa)
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionBa + ":" + this.ProportionBa;

		if (!IsNullOrEmpty(this.ProportionBaMin))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionBaMin + ":" + this.ProportionBaMin;

		if (!IsNullOrEmpty(this.ProportionBaMax))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionBaMax + ":" + this.ProportionBaMax;

		if (!this.ProportionCs)
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionCs + ":" + this.ProportionCs;

		if (!IsNullOrEmpty(this.ProportionCsMin))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionCsMin + ":" + this.ProportionCsMin;

		if (!IsNullOrEmpty(this.ProportionCsMax))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ProportionCsMax + ":" + this.ProportionCsMax;

		if (!IsNullOrEmpty(this.LongTerm))
			retHash += (retHash != "" ? ";" : "") + this.Keys.LongTerm + ":" + this.LongTerm;

		if (!IsNullOrEmpty(this.ShortTerm))
			retHash += (retHash != "" ? ";" : "") + this.Keys.ShortTerm + ":" + this.ShortTerm;

		if (!IsNullOrEmpty(this.CompanyId))
			retHash += (retHash != "" ? ";" : "") + this.Keys.CompanyId + ":" + this.CompanyId;

		if (!IsNullOrEmpty(this.GetCompanyProjectsOnly))
			retHash += (retHash != "" ? ";" : "") + this.Keys.GetCompanyProjectsOnly + ":" + this.GetCompanyProjectsOnly;

		if (!IsNullOrEmpty(this.IncludeNonAwarded))
			retHash += (retHash != "" ? ";" : "") + this.Keys.IncludeNonAwarded + ":" + this.IncludeNonAwarded;

		if (!IsNullOrEmpty(this.SearchAllCompanies))
			retHash += (retHash != "" ? ";" : "") + this.Keys.SearchAllCompanies + ":" + this.SearchAllCompanies;

		if (!IsNullOrEmpty(retHash)) retHash = "#" + retHash;
		return retHash;
	}
}


function URLParamKeys() {
	this.CurrentPage = "p";
	this.ProjectType = "pt";
	this.Keyword = "k";
	this.Donor = "d";
	this.DonorId = "did";
	this.Country = "c";
	this.Region = "rg";
	this.SectorList = "sc";
	this.MainSector = "msc";
	this.Status = "st";
	this.Unit = "u";
	this.BudgetMin = "bmin";
	this.BudgetMax = "bmax";
	this.Orderby = "ord";
	this.YearFrom = "yfr";
	this.YearTo = "yto";
	this.ProportionWd = "pw";
	this.ProportionWdMin = "pwmn";
	this.ProportionWdMax = "pwmx";
	this.ProportionBa = "pb";
	this.ProportionBaMin = "pbmn";
	this.ProportionBaMax = "pbmx";
	this.ProportionCs = "pc";
	this.ProportionCsMin = "pcmn";
	this.ProportionCsMax = "pcmx";
	this.LongTerm = "lter";
	this.ShortTerm = "ster";
	this.CompanyId = "cid";
	this.GetCompanyProjectsOnly = "cponly";
	this.IncludeNonAwarded = "inaw";
	this.SearchAllCompanies = "sac";
}


var urlParams = new URLParams();


function UpdateUrlHash() {
	// update location hash (for "back" button compatibility):
	var paramsLocation = urlParams.assembleUrlHash();
	$("paramsHolder").attr("name", paramsLocation.replace("#", ""));
	document.location = paramsLocation;
}