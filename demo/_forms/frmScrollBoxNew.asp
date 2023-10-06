<% 
Dim iCouScrllHeight, iDonScrllHeight, iSctScrllHeight
Dim full_shift, is_mozilla, is_nav, is_chrome
Dim bShowAllDonors
Dim iMaxSectors
' For experts' SIP and member's Free Trial BSC - limitation on 80 sectors 
' SIP updates from IBF's network are enabled for unlimited number of sectors
If (sScriptFileName="sc_register.asp") Or (sScriptFileName="sc_update.asp" And Left(sUserIpAddress, 10)<>"158.29.157") Or (sScriptFileName="bsc_register.asp" And (sAccessType="trial" Or sAccessType="cml")) Then
	iMaxSectors=80
Else
	iMaxSectors=400
End If


' to identify later
Dim ts, tt, k1, k2, mreg, mExF, m, j, mExT, mExTInt, strTemp, cp, mGZ, mNt, mNtInt, mOrg, mOrgInt

If aNt<12 Then
	iCouScrllHeight=60+21*aNt
Else
	iCouScrllHeight=311
End If

If aOrg<=9 Then
	iDonScrllHeight=26+21*aOrg
Else
	iDonScrllHeight=225
End If

If aExT<12 Then
	iSctScrllHeight=60+21*aExT+full_shift*13
Else
	iSctScrllHeight=352
End If


sUserAgent=LCase(Request.ServerVariables("HTTP_USER_AGENT"))
If (InStr(sUserAgent,"mozilla")>0) AND (InStr(sUserAgent,"spoofer")=0) AND (InStr(sUserAgent,"compatible")=0) AND (InStr(sUserAgent,"opera555")=0) AND (InStr(sUserAgent,"webtv")=0) AND (InStr(sUserAgent,"gecko")=0) Then
	is_nav = 1
Else
	is_nav = 0
End If
If (InStr(sUserAgent,"mozilla")>0) AND (InStr(sUserAgent,"gecko")>0) Then
	is_mozilla = 1
Else
	is_mozilla = 0
End If
If (InStr(sUserAgent,"chrome")>0) Then
	is_chrome = 1
	is_mozilla = 0
Else
	is_chrome = 0
End If

Sub InsertScrollStyles
	
End Sub


Sub InsertJSScrollFunctions(bShowTotal, bShowAll)
%>
<script language="JavaScript" type="text/javascript">
<!--

// **********************************************
function RoundTo5(amount)
{
var addamount, newamount, lastdigit;
newamount=Math.round(amount);
newamountstring=''+newamount
lastdigit=newamountstring.substr(newamountstring.length-1, 1);
if (lastdigit==0) {addamount=0;}
if (lastdigit==1) {addamount=4;}
if (lastdigit==2) {addamount=3;}
if (lastdigit==3) {addamount=2;}
if (lastdigit==4) {addamount=1;}
if (lastdigit==5) {addamount=0;}
if (lastdigit==6) {addamount=4;}
if (lastdigit==7) {addamount=3;}
if (lastdigit==8) {addamount=2;}
if (lastdigit==9) {addamount=1;}

return(newamount+addamount);
}

// **********************************************
function GetControl(layerName){
  if(document.getElementById) return document.getElementById(layerName)
  if(document.all) return document.all[layerName]
  if(document.layers) return eval('document.' + layerName)
}

// **********************************************
function ChangeControlColor(lControl, lItemNum, lColorNum)
{
var tColor, tControl;
  tControl=GetControl(lControl + 'Text' + lItemNum.toString());
  if(lColorNum==1) {tColor='#CC0000';} else {tColor='#000066';}
  tControl.style.color=tColor;
  //tControl.color=tColor;
}

// **********************************************
function ChangeControlText(lControl, lItemNum, lItemValue )
{
var tControl, tPos;
if (is_ie4up || is_mozilla) 
{
  <% If is_nav<>1 Then %>
  tControl=GetControl(lControl + 'Text' + lItemNum.toString());
  tPos=tControl.innerHTML.indexOf('(');
  if(lItemValue>0)
	{if(tPos>0)
	{tControl.innerHTML=tControl.innerHTML.substring(0,tPos-1);}
	tControl.innerHTML=tControl.innerHTML + ' (' + lItemValue.toString() + ')';
	  if(lItemValue>0)
	  {ChangeControlColor(lControl, lItemNum, 1);}
	}
  else
	{tControl.innerHTML=tControl.innerHTML.substring(0,tPos-1);
	ChangeControlColor(lControl, lItemNum, 0);}
  <% End If %>
}
}





// **********************************************

// **********************************************
<% If bShowTotal=1 Then %>
function SetTotal(cfield)
{                 
<% If sScriptFileName<>"update_prf2.asp" And sScriptFileName<>"register_prf2.asp" Then %>
var cntTotal = document.RegForm.mmb_total_price.value;
var activeCurrency;
var activeExchangeRate;
activeExchangeRate=<%=Replace(ExchangeRate("EUR", "USD", Now()), ",", ".")%>;

  document.RegFormCou.mmb_cou_total.value=mmb_cou;
  document.RegFormDon.mmb_don_total.value=mmb_don;
  document.RegFormSct.mmb_sct_total.value=mmb_sct;

    <% If sAccessType<>"trial"  Or sAccessType<>"cml" Then %>
   if (cfield=='cur2')
   { document.RegForm.mmb_total_currency1.selectedIndex=document.RegForm.mmb_total_currency2.selectedIndex; }
   else if (cfield=='cur1')
   { document.RegForm.mmb_total_currency2.selectedIndex=document.RegForm.mmb_total_currency1.selectedIndex; }
    <% End If %>
  
  if (mmb_sct>0 && mmb_cou>0 && mmb_don>0)
  {
    if ((mmb_sct>0) && (mmb_sct<21))  {cntSct=1;}
    if ((mmb_sct>20) && (mmb_sct<41)) {cntSct=2;}
    if ((mmb_sct>40) && (mmb_sct<81)) {cntSct=3;}
    if ((mmb_sct>80) && (mmb_sct<161)){cntSct=4;}
    if (mmb_sct>160) { cntSct=5 }

    if ((mmb_cou>0) && (mmb_cou<16))  {cntCou=1;}
    if ((mmb_cou>15) && (mmb_cou<41)) {cntCou=2;}
    if ((mmb_cou>40) && (mmb_cou<81)) {cntCou=3;}
    if ((mmb_cou>80) && (mmb_cou<121)){cntCou=4;}
    if (mmb_cou>120) { cntCou=5 }

    cntDon=mmb_don;
    // ignoring additional funding agencies
    for (i=1; i<=<%=aOrg%>; i++)
    { if (jOrgInt[i]>0 && jOrgMain[i]==0) {--cntDon;}}

    if ((cntDon>0) && (cntDon<3)) {cntDon=1;}
    if ((cntDon>2) && (cntDon<5)) {cntDon=2;}
    if ((cntDon>4) && (cntDon<7)) {cntDon=3;}
    if ((cntDon>6) && (cntDon<20)) {cntDon=3;}

  <% If Not iExpertID>0 Then %>
    <% If sAccessType="trial" Or sAccessType="cml" Then %>
    <% Else %>
    for (var i=0;i<4;i++){
		if (document.RegForm.mmb_period[i] && document.RegForm.mmb_period[i].checked == true) {
			document.RegForm.mmb_period_hid.value=document.RegForm.mmb_period[i].value;
		}
	}

	if (document.RegForm.mmb_period && document.RegForm.mmb_period.checked == true) {
		document.RegForm.mmb_period_hid.value=document.RegForm.mmb_period.value;
	}

    // 90 eur - added subscription fee for devbusiness
    cntTotal= jPrice[cntSct][cntCou][cntDon];

    if (cntTotal>0)
    {activeCurrency=(document.RegForm.mmb_total_currency1.options[document.RegForm.mmb_total_currency1.selectedIndex].value=="EUR")?1:activeExchangeRate;
     document.RegForm.mmb_total_price.value=RoundTo5(cntTotal*activeCurrency);
     document.RegForm.mmb_total_sum.value=document.RegForm.mmb_total_price.value*document.RegForm.mmb_period_hid.value;}
    <% End If %>

  <% End If %>  
  }
  else
  {document.RegForm.mmb_total_price.value='';
  document.RegForm.mmb_total_sum.value='';}
<% End If %>
}
<% End If %>




function LoadInt()
{

if (document.RegForm.mmb_cou_hid.value.length>5)
{jNtInt = document.RegForm.mmb_cou_hid.value.split(',');
if (jNtInt[0]>'')
{mmb_cou=jNtInt[0];jNtInt[0]='';
<% If bShowTotal=1 Then %>
document.RegFormCou.mmb_cou_total.value=mmb_cou;
<% End If %>
}}

if (document.RegForm.mmb_don_hid.value.length>5)
{jOrgInt = document.RegForm.mmb_don_hid.value.split(',');
if (jOrgInt[0]>'')
{mmb_don=jOrgInt[0];jOrgInt[0]='';
<% If bShowTotal=1 Then %>
document.RegFormDon.mmb_don_total.value=mmb_don;
<% End If %>
}}

if (document.RegForm.mmb_sct_hid.value.length>5)
{jExTInt = document.RegForm.mmb_sct_hid.value.split(',');
if (jExTInt[0]>'')
{mmb_sct=jExTInt[0];jExTInt[0]='';
<% If bShowTotal=1 Then %>
document.RegFormSct.mmb_sct_total.value=mmb_sct;
<% End If %>
}}

<% If bShowTotal=1 Then %>
SetTotal();
<% End If %>
}

// **********************************************
function ShowPrice(sSender) 
{  
  var params, pfile;
  if (sSender=='bsc_reg')
  {
	if (document.RegForm.mmb_total_currency1.options[document.RegForm.mmb_total_currency1.selectedIndex].value=="USD")
	{pfile='bsc_price.asp?dcr=USD'}
	else
	{pfile='bsc_price.asp?dcr=EUR'}
	params=(cntCou>0 && cntDon>0 && cntSct>0)? '?prm=1&cou='+cntCou+'&don='+cntDon+'&sct='+cntSct : '';
  }
  window.open(pfile+params,'ANWnd','scrollbars=yes,status=yes,resizable=yes,menubar=yes');
}

// -->
</script>
<%
End Sub

Sub InsertJsHelpers(bShowTotal, bShowAll)

	%><script language="javascript" type="text/javascript">

	jNtInt = new Array();
	jNtCode= new Array();
	jNtZone= new Array();
	jGZnInt = new Array();
	jOrgInt = new Array();
	jOrgCode = new Array();
	jOrgMain = new Array();
	jMDonInt = new Array();
	jExTInt= new Array();
	jExTCode= new Array();
	jExTSrch= new Array();
	jExFInt = new Array();

	var rst_cou=new Array();
	var rst_don=new Array();
	var rst_sct=new Array();
	var agt;

	var cntCou=0;
	var cntDon=0;
	var cntSct=0;

	Chk1 = new Image(13,13); Chk1.src = 'n.gif';
	Chk2 = new Image(13,13); Chk2.src = 'c.gif';

	// COUNTRIES:
	function RReg(regId)
	{
		var isChecked = $("#reg_" + regId).is(":checked");

		// check/uncheck countries within a region, or all, if a region Id was not specified(=0):
		$("#divCouSelector input[type='checkbox'][id^='cou_" + (regId == 0 ? "" : (regId + "_")) + "']").each(function () {
			$(this).prop("checked", isChecked);
		});

		if (regId == 0) {
			// check/uncheck all regions:
			$("#divCouSelector input[type='checkbox'][id^='reg_']").each(function () {

				$(this).prop("checked", isChecked);

				// update all regions totals:
				var regionId = $(this).attr("id").replace('reg_', '');
				var regTotal = CountCheckedBySelector("#divCouSelector input[type='checkbox'][id^='cou_" + regionId + "_']");
				$("#RegTotal_" + regionId).html(regTotal > 0 ? ("(" + regTotal + ")") : "");
				$("#Reg_" + regId).attr("class", (regTotal > 0 ? "red" : ""));
			});
		}
		else {
			// if a region is unchecked - incheck also the "select all" checkbox:
			$(this).prop("checked", isChecked);

			// update single region total:
			var regTotal = CountCheckedBySelector("#divCouSelector input[id^='cou_" + regId + "_']");
			$("#RegTotal_" + regId).html(regTotal > 0 ? ("(" + regTotal + ")") : "");
			$("#Reg_" + regId).attr("class", (regTotal > 0 ? "red" : ""));
		}

		// To optimise the performance in IE the hidden field will be updated on form submit
		// UpdateHidFldForScrollbox("mmb_cou_hid", "divCouSelector", "cou_");
		
		// update the Main total:
		/*
		var allTotal = CountCheckedBySelector("#divCouSelector input[type='checkbox'][id^='cou_']");
		$("#Cou_AllTotal").html(allTotal > 0 ? ("Total : " + allTotal) : "");
		*/
	}

	function RCou(regId, couId)
	{
		var isChecked = $("#cou_" + regId + "_" + couId).is(":checked");

		// update region total:
		var regTotal = CountCheckedBySelector("#divCouSelector input[id^='cou_" + regId + "_']");
		$("#RegTotal_" + regId).html(regTotal > 0 ? ("(" + regTotal + ")") : "");
		$("#Reg_" + regId).attr("class", (regTotal > 0 ? "red" : ""));

		// unckeck the region and the "select all" checkbox if any country unchecked:
		if (!isChecked) {
			$("#reg_" + regId).prop("checked", isChecked);
			$("#reg_0").prop("checked", isChecked);
		}

		// To optimise the performance in IE the hidden field will be updated on form submit
		// UpdateHidFldForScrollbox("mmb_cou_hid", "divCouSelector", "cou_");

		// update the Main total:
		/*
		var allTotal = CountCheckedBySelector("#divCouSelector input[type='checkbox'][id^='cou_']");
		$("#Cou_AllTotal").html(allTotal > 0 ? ("Total : " + allTotal) : "");
		*/
	}

	// DONORS:
	function RDon(grpId, donId) {

		if (grpId == 0) {
			var isChecked = $("#don_" + grpId).is(":checked");

			// check/uncheck all donors:
			$("#divDonSelector input[id^='don_']").each(function () {
				$(this).prop("checked", isChecked);
			});

			$("#divDonSelector a[id^='don_']").each(function () {
				// update all group totals:
				var groupId = $(this).attr("id").replace('don_', '');
				var grpTotal = CountCheckedBySelector("#divDonSelector input[id^='don_" + groupId + "_']");
				$("#DonTotal_" + groupId).html(grpTotal > 0 ? ("(" + grpTotal + ")") : "");
			});
		}
		else {
			var isChecked = $("#don_" + grpId + "_" + donId).is(":checked");

			// update group total:
			var grpTotal = CountCheckedBySelector("#divDonSelector input[id^='don_" + grpId + "_']");
			$("#DonTotal_" + grpId).html(grpTotal > 0 ? ("(" + grpTotal + ")") : "");

			// uncheck the group and the "select all" checkbox if any donor unchecked:
			if (!isChecked) {
				$("#don_" + grpId).prop("checked", isChecked);
				$("#don_0").prop("checked", isChecked);
			}
		}

		// To optimise the performance in IE the hidden field will be updated on form submit
		// UpdateHidFldForScrollbox("mmb_don_hid", "divDonSelector", "don_");

		// update the Main total:
		/*
		var allTotal = CountCheckedBySelector("#divDonSelector input[id^='don_']");
		$("#Don_AllTotal").html(allTotal > 0 ? ("Total : " + allTotal) : "");
		*/
	}

	// SECTORS:

	function RMSct(msctId) {
		var isChecked = $("#msct_" + msctId).is(":checked");

		// check/uncheck sectors within a main sector, or all, if a main sector Id was not specified(=0):
		$("#divSctSelector input[type='checkbox'][id^='sct_" + (msctId == 0 ? "" : (msctId + "_")) + "']").each(function () {
			$(this).prop("checked", isChecked);
		});

		if (msctId == 0) {
			// check/uncheck all main sectors:
			$("#divSctSelector input[type='checkbox'][id^='msct_']").each(function () {
				$(this).prop("checked", isChecked);
				// update all main sector totals:
				var mainsectorId = $(this).attr("id").replace('msct_', '');
				var msctTotal = CountCheckedBySelector("#divSctSelector input[id^='sct_" + mainsectorId + "_']");
				$("#MSctTotal_" + mainsectorId).html(msctTotal > 0 ? ("(" + msctTotal + ")") : "");
				$("#MSct_" + mainsectorId).attr("class", (msctTotal > 0 ? "red" : ""));
			});
		}
		else {
			// if a main sector is unchecked - uncheck also the "select all" checkbox:
			if (!isChecked) {
				$("#msct_0").prop("checked", isChecked);
			}

			// update single main sector total:
			var msctTotal = CountCheckedBySelector("#divSctSelector input[id^='sct_" + msctId + "_']");
			$("#MSctTotal_" + msctId).html(msctTotal > 0 ? ("(" + msctTotal + ")") : "");
			$("#MSct_" + msctId).attr("class", (msctTotal > 0 ? "red" : ""));
		}

		// To optimise the performance in IE the hidden field will be updated on form submit
		// UpdateHidFldForScrollbox("mmb_sct_hid", "divSctSelector", "sct_");

		// update the Main total:
		/*
		var allTotal = CountCheckedBySelector("#divSctSelector input[id^='sct_']");
		$("#Sct_AllTotal").html(allTotal > 0 ? ("Total : " + allTotal) : "");
		*/
	}

	function RSct(msctId, sctId) {
		var isChecked = $("#sct_" + msctId + "_" + sctId).is(":checked");

		// update main sector total:
		var msctTotal = CountCheckedBySelector("#divSctSelector input[id^='sct_" + msctId + "_']");
		$("#MSctTotal_" + msctId).html(msctTotal > 0 ? ("(" + msctTotal + ")") : "");
		$("#MSct_" + msctId).attr("class", (msctTotal > 0 ? "red" : ""));

		// uncheck the main sector and the "select all" checkbox if any sector unchecked:
		if (!isChecked) {
			$("#msct_" + msctId).prop("checked", isChecked);
			$("#msct_0").prop("checked", isChecked);
		}

		// To optimise the performance in IE the hidden field will be updated on form submit
		// UpdateHidFldForScrollbox("mmb_sct_hid", "divSctSelector", "sct_");

		// update the Main total:
		/*
		var allTotal = CountCheckedBySelector("#divSctSelector input[type='checkbox'][id^='sct_']");
		$("#Sct_AllTotal").html(allTotal > 0 ? ("Total : " + allTotal) : "");
		*/
	}

	// GENERICS:

	function ScrollToBlock(headerId, containerId) {
		var headerObj = $("#header_" + headerId)[0];
		var containerDiv = $("#" + containerId)[0];
		$("#" + containerId).mCustomScrollbar("scrollTo", headerObj.offsetTop);
	}

	function CountCheckedBySelector(selector) {
		var selectedCnt = 0;

		$(selector).each(function () {
			if ($(this).is(":checked")) selectedCnt++;
		});

		return selectedCnt;
	}

	function GetScrollboxSelection(scrollBoxId, cbIdPrefix) {
		var csVals = "";
		// get all checkboxes, which start with the specified checkbox Id prefix and get their values as a string:
		$("#" + scrollBoxId + " input[type='checkbox'][id^='" + cbIdPrefix + "']").each(function () {
			if ($(this).is(":checked") && $(this).val() != "")
				csVals += (csVals != '' ? ',' : '') + $(this).val();
		});
		return csVals;
	}

	function UpdateHidFldForScrollbox(hidFldId, scrollBoxId, cbIdPrefix) {
		var csVals = GetScrollboxSelection(scrollBoxId, cbIdPrefix);
		$("#" + hidFldId).val(csVals);
	}

	// OTHER:
	function RestoreInt() {
		// IE doesn't restore checkboxes automatically
		var f = document.RegForm;

		var countries;
		if (f && f.mmb_cou_hid && f.mmb_cou_hid.value) {
			countries  = f.mmb_cou_hid.value.split(",");
			for (var i = 0, max = countries.length; i < max; i++) {
				var reg = cou_reg[countries[i]];
				if (reg) {
					//console.log("cou_" + reg + "_" + countries[i]);
					$("#cou_" + reg + "_" + countries[i]).prop("checked", true);
				}
			}

			// set region totals:
			$('#divCouList input[type="checkbox"][id^="reg_"]').each(function () {
				var regId = $(this).attr("id").replace("reg_", "");
				// update region total:
				var regTotal = CountCheckedBySelector("#divCouList input[type='checkbox'][id^='cou_" + regId + "_']");
				$("#RegTotal_" + regId).html(regTotal > 0 ? ("(" + regTotal + ")") : "");
				$("#Reg_" + regId).attr("class", (regTotal > 0 ? "red" : ""));
			});
		}

		var sectors;
		if (f && f.mmb_sct_hid && f.mmb_sct_hid.value) {
			sectors  = f.mmb_sct_hid.value.split(",");
			for (var i = 0, max = sectors.length; i < max; i++) {
				var msct = sct_msct[sectors[i]];
				if (msct) {
					//console.log("sct_" + msct + "_" + sectors[i]);
					$("#sct_" + msct + "_" + sectors[i]).prop("checked", true);
				}
			}

			// set main sectors totals:
			$('#divSctList input[type="checkbox"][id^="msct_"]').each(function () {
				var msctId = $(this).attr("id").replace("msct_", "");
				// update main sector total:
				var msctTotal = CountCheckedBySelector("#divSctList input[type='checkbox'][id^='sct_" + msctId + "_']");
				$("#MSctTotal_" + msctId).html(msctTotal > 0 ? ("(" + msctTotal + ")") : "");
				$("#MSct_" + msctId).attr("class", (msctTotal > 0 ? "red" : ""));
			});
		}


		var donors;
		if (f && f.mmb_don_hid && f.mmb_don_hid.value) {
			donors  = f.mmb_don_hid.value.split(",");
			for (var i = 0, max = donors.length; i < max; i++) {
				var dmain = don_main[donors[i]];
				if (dmain) {
					var dmain2 = 2 - dmain;
					//console.log("don" + dmain + "_" + donors[i]);
					$("#don_" + dmain2 + "_" + donors[i]).prop("checked", true);
				}
			}

			// set donor groups totals:
			for (i = 1; i <= 2; i++) {
				var grpTotal = CountCheckedBySelector("#divDonList input[type='checkbox'][id^='don_" + i + "_']");
				$("#DonTotal_" + i).html(grpTotal > 0 ? ("(" + grpTotal + ")") : "");
				$("#Don_" + i).attr("class", (grpTotal > 0 ? "red" : ""));
			}
		}
	}
	</script>
	<%
End Sub

' COUNTRIES selector:
Sub ShowCouScrollBox(sBoxTitle, sLeftMenuTitle, bShowLeftMenu, bShowAll, bShowRegions, bShowCountries, bShowTotal)
	%>
	<div id="divCouSelector" class="filter_table">
		<% ' title:
		If sBoxTitle > "" Then Response.Write("<div class=""filter_table_header""><h3>" & sBoxTitle & "</h3></div>")
		
		' LEFT COLUMN: 
		if bShowLeftMenu = 1 Then
			%><div class="filter_table_filters">
				<% ' title above the left section, if supplied:
				If sLeftMenuTitle > "" Then Response.Write("<p class=""filter_table_filters_header"">" & sLeftMenuTitle & "</p>")
				%>
				<ul><% 
				Set objTempRs = objConn.Execute("EXEC usp_DatContinentSelect")
				j = 0
				While Not objTempRs.Eof
					FOR i = 0 to aGZ - 1 
						If aGZnContinent(i) = objTempRs("id_Continent") Then 
							%><a href="javascript:void(0);" onclick="ScrollToBlock('cou_<%=aGZnCode(i)%>', 'divCouList')"><li><p><span id="Reg_<%=aGZnCode(i)%>"><%=aGZnInfo(i)%></span> <span id="RegTotal_<%=aGZnCode(i)%>" class="filter_table_nbr_selected red"></span></p></li></a>
							<%
						End If
					Next
					objTempRs.MoveNext
					j = j + 1
				WEnd 
				objTempRs.Close
				Set objTempRs=Nothing 

				' main total:
				If bShowTotal = 1 Then 
					%><a href="javascript:void(0)"><li class="last"><p id="Cou_AllTotal" class="red bold"></p></li></a>
					<%
				End If 
				%>
				</ul>
			</div>
			<%
		End If

		' RIGHT COLUMN: %>
		<div id="divCouList" class="filter_table_object_wrapper">
		<%
		' "select all countries" checkbox:
		If bShowAll = 1 Then 
			%><div class="filter_table_main_obj"><label class="filter_table_checkbox"><input type="checkbox" name="reg_0" id="reg_0" onclick="RReg(0)" value="" /><!--span></span--></label><label class="option_label_text" for="reg_0"><span>Select all countries</span></label></div>
			<%
		End If
		%>
		<%
		FOR i = 0 to aGZ - 1 
			' "select region" checkbox:
			%>
			<div id="header_cou_<%=aGZnCode(i)%>" class="filter_table_main_obj">
				<label class="filter_table_checkbox">
					<input type="checkbox" id="reg_<%=aGZnCode(i)%>" name="reg_<%=aGZnCode(i)%>" onclick="RReg(<%=aGZnCode(i)%>)" /><!--span></span-->
				</label>
				<label class="option_label_text" for="reg_<%=aGZnCode(i)%>"><span><%=aGZnInfo(i)%></span><br/>[select all countries in the region]</label>
			</div>
			<ul>
			<%
			' aGZnInfo(i) - geozone name
			' aGZnCode(i) - geozone ID
			' aGZnContinent(i) = geozone continent
			
			' aNtInfo(i) - country name
			' aNtCode(i) - country ID
			' aNtZone(i) - geozone ID for the country
			For j = 0 To aNt - 1
				If aNtZone(j) = aGZnCode(i) Then
					mreg = 0
					For m = 0 To mNt-1 
						If aNtCode(j) = mNtCode(m) Then 
							mreg = mNtCode(m) 
							Exit For 
						End If 
					Next

					If mreg > 0 Then mNtInt = mNtInt & "," & mreg 

					' "select county" checkbox:
					%><li>
						<label class="filter_table_checkbox">
							<input type="checkbox" id="cou_<%=aGZnCode(i)%>_<%=aNtCode(j)%>" name="cou_<%=aGZnCode(i)%>_<%=aNtCode(j)%>" value="<%=aNtCode(j)%>" onclick="RCou(<%=aGZnCode(i)%>, <%=aNtCode(j)%>);" <% If mreg > 0 Then %>checked="checked"<% End If %>/><!--span></span-->
						</label>
						<label class="option_label_text" for="cou_<%=aGZnCode(i)%>_<%=aNtCode(j)%>"><%=aNtInfo(j)%></label>
					</li>
					<% 
				End If 
			Next 
			%>
			</ul>
			<%
		Next
		%>
		</div>
	</div>
	<%
End Sub

' DONORS selector:
Sub ShowDonScrollBox(sBoxTitle, sLeftMenuTitle, bShowLeftMenu, bShowAll, bShowTotal, bDonorsDelimited)
	%>
	<div id="divDonSelector" class="filter_table">
		<% 
		If sBoxTitle > "" Then Response.Write("<div class=""filter_table_header""><h3>" & sBoxTitle & "</h3></div>") 

		' LEFT COLUMN: 
		if bShowLeftMenu = 1 Then
			%>
			<div class="filter_table_filters"><%
				If sLeftMenuTitle>"" Then Response.Write("<p class=""filter_table_filters_header"">" & sLeftMenuTitle & "</p>")
				%>
				<ul>
					<a href="javascript:void(0);" id="don_1" onclick="ScrollToBlock('don_1', 'divDonList')"><li><p>Major funding agencies <span id="DonTotal_1" class="filter_table_nbr_selected red"></span></p></li></a>
					<a href="javascript:void(0);" id="don_2" onclick="ScrollToBlock('don_2', 'divDonList')"><li><p>Bonus funding agencies <span id="DonTotal_2" class="filter_table_nbr_selected red"></span></p></li></a>
				<%
				' main total:
				If bShowTotal = 1 Then 
					%><a href="javascript:void(0)"><li class="last"><p id="Don_AllTotal" class="red bold"></p></li></a>
					<%
				End If 
				%>
				</ul>
			</div>
			<%
		End If

		' RIGHT COLUMN: %>
		<div id="divDonList" class="filter_table_object_wrapper">
		<%
		If bShowAll = 1 Then 
			%><div class="filter_table_main_obj"><label class="filter_table_checkbox"><input type="checkbox" name="don_0" id="don_0" onclick="RDon(0, 0)" value="" /><!--span></span--></label><label class="option_label_text" for="don_0"><span>Select all funding agencies</span></label></div>
			<% 
		End If
		%>
		<ul>
		<% ' major agencies: %>
		<div id="header_don_1" style="height:1px;"></div>
		<%  
		Dim donType ' 1 = main ; 2 = bonus
		donType = 1
		FOR j = 0 to aOrg - 1
			If aOrgMainDonor(j) = 0 Then
				donType = 2
			Else
				donType = 1
			End If

			mreg = 0
			FOR m = 0 to mOrg - 1 
				If aOrgCode(j) = mOrgCode(m) Then 
					mreg = mOrgCode(m) 
					Exit For 
				End If 
			NEXT

			If mreg > 0 Then mOrgInt = mOrgInt & "," & mreg

			If aOrgMainDonor(j) = 0 And bDonorsDelimited = 0 Then 
				bDonorsDelimited = 1
				' bonus agencies:
				%><%
				If sScriptFileName = "register_bsc.asp" Or sScriptFileName = "update_bsc.asp" Then 
					%><div id="header_don_2" style="height:2px; background-color: #ECECEC; color: #1A466A; margin: 5px 0;"></div>
						<label class="filter_table_checkbox"></label>
						<label class="option_label_text"><span>Free bonus funding agencies</span></label>
					</div>
					<%
				else
					%><div id="header_don_2" style="height:2px; background-color: #ECECEC; color: #1A466A; margin: 5px 0;"></div>
					<%
				End If 
				%>
				<%
			End If
			
			' "select donor" checkbox:
			%><li>
				<label class="filter_table_checkbox">
					<input type="checkbox" name='don_<%=donType %>_<%=aOrgCode(j)%>' id="don_<%=donType %>_<%=aOrgCode(j)%>" value="<%=aOrgCode(j)%>" onclick="RDon(<%=donType %>, <%=aOrgCode(j)%>);" <% If mreg > 0 Then %>checked="checked"<% End If %>/><!--span></span-->
				</label>
				<label class="option_label_text" for="don_<%=donType %>_<%=aOrgCode(j)%>"><%=aOrgInfo(j)%></label>
			</li>
			<% 
		NEXT %>
		</ul>
		</div>
	</div>
	<%
End Sub

' SECTORS selector:
Sub ShowSctScrollBox(sBoxTitle, sLeftMenuTitle, bShowLeftMenu, bShowAll, bShowTotal)
	%>
	<div id="divSctSelector" class="filter_table">
		<% 
		If sBoxTitle > "" Then Response.Write("<div class=""filter_table_header""><h3>" & sBoxTitle & "</h3></div>")
		
		' LEFT COLUMN: 
		if bShowLeftMenu = 1 Then
			%><div class="filter_table_filters">
				<% ' title above the left section, if supplied:
				If sLeftMenuTitle > "" Then Response.Write("<p class=""filter_table_filters_header"">" & sLeftMenuTitle & "</p>")
				%>
				<ul><%
				For i = 0 To aExF - 1
					%><a href="javascript:void(0);" onclick="ScrollToBlock('sct_<%=aExFCode(i)%>', 'divSctList')"><li><p><span id="MSct_<%=aExFCode(i)%>"><%=aExFShort(i)%></span> <span id="MSctTotal_<%=aExFCode(i)%>" class="filter_table_nbr_selected red"></span></p></li></a>
					<% 
				Next 

				' main total:
				If bShowTotal = 1 Then 
					%><a href="javascript:void(0)"><li class="last"><p id="Sct_AllTotal" class="red bold"></p></li></a>
					<%
				End If 
				%>
				</ul>
			</div>
			<%
		End If

		' RIGHT COLUMN: %>
		<div id="divSctList" class="filter_table_object_wrapper">
		<%' "select all sectors" checkbox:
		If bShowAll = 1 Then
			%><div class="filter_table_main_obj"><label class="filter_table_checkbox"><input type="checkbox" name="msct_0" id="msct_0" onclick="RMSct(0)" /><!--span></span--></label><label class="option_label_text" for="msct_0"><span>Select all sectors and sub-sectors</span></label></div>
			<%
		End If
		
		%>
		<%
		For i = 0 To aExF - 1
			' "select main sector" checkbox:
			%><div id="header_sct_<%=aExFCode(i)%>" class="filter_table_main_obj">
				<label class="filter_table_checkbox">
					<input type="checkbox" id="msct_<%=aExFCode(i)%>" name="msct_<%=aExFCode(i)%>" onclick="RMSct(<%=aExFCode(i)%>)" /><!--span></span-->
				</label>
				<label class="option_label_text" for="msct_<%=aExFCode(i)%>"><span><% =CutStringNSplit(aExFInfo(i), 52, "<br>") %></span><br/>[select all sub-sectors]</label>
			</div>
			<ul>
			<%
			For j = 0 To aExT - 1 
				If aExTSrch(j) = aExFCode(i) Then
					mreg = 0
					For m = 0 To mExT-1
						If aExTCode(j) = mExTCode(m) Then
							mreg = mExTCode(m)
							Exit For
						End If
					Next
					
					If mreg > 0 Then mExTInt = mExTInt & "," & mreg

					' "select sub-sector" checkbox:
					%><li>
						<label class="filter_table_checkbox">
							<input type="checkbox" id="sct_<%=aExFCode(i)%>_<%=aExTCode(j)%>" name="sct_<%=aExFCode(i)%>_<%=aExTCode(j)%>" value="<%=aExTCode(j)%>" onclick="RSct(<%=aExFCode(i)%>, <%=aExTCode(j)%>);" <% If mreg > 0 Then %>checked="checked"<% End If %>/><!--span></span-->
						</label>
						<label class="option_label_text" for="cou_<%=aExFCode(i)%>_<%=aExTCode(j)%>"><% =CutStringNSplit(aExTInfo(j), 55, "<br>") %></label>
					</li>
					<%
				End If
			Next
			%>
			</ul>
			<%
		Next
		%>
		</div>
	</div>
	<% If InStr(sScriptFileName, "exp_search") > 0 Then %>
	<div class="filter_table_flex">
		<div style="margin: 5px 5px 5px 195px;">
			<label class="filter_table_checkbox">
				&nbsp;&nbsp;<input type="checkbox" name="sectors_simultaneously" /><!--span></span-->
			</label>
			<label class="option_label_text" for="sectors_simultaneously">All selected sectors simultaneously</label>
		</div>
	</div>
	<% End If %>
	<% 
End Sub

Sub WriteFilterTableScript
%>
<!-- Script for filter table -->
<script language="javascript">
	//This script use to ajust "filter_table_object_wrapper" to have the same height than "filter_table_filters"
	function matchHeight() {
		$('.filter_table').each(function () {
			var newHeight = $(this).children(".filter_table_filters").height(); //where #grownDiv is what's growing
			//	alert($(".filter_table_filters").attr('id') + ' ' + newHeight);
			if (newHeight < 400) newHeight = 400;

			$(this).children(".filter_table_object_wrapper").height(newHeight);    //where .matchDiv is the class of the other two
		});
	}
	jQuery.event.add(window, "load", matchHeight);
	//This script is the call for the customscrollbar
	(function ($) {
		$(window).load(function () {
			$(".filter_table_object_wrapper").mCustomScrollbar(
            {
            	theme: "dark-thick",
            	autoHideScrollbar: true,
            	scrollInertia: 0
            });
		});
	})(jQuery);

	var cou_reg = {};
	<% For i = 0 To aNt - 1 %>
	cou_reg["<% =aNtCode(i) %>"] = "<% =aNtZone(i) %>"; <% Next %>

	var don_main = {};
	<% For i = 0 To aOrg - 1 %>
	don_main["<% =aOrgCode(i) %>"] = "<% =aOrgMainDonor(i) %>"; <% Next %>

	var sct_msct = {};
	<% For i = 0 To aExT - 1 %>
	sct_msct["<% =aExTCode(i) %>"] = "<% =aExTSrch(i) %>"; <% Next %>
</script>
<!-- End Script for filter table -->
<%
End Sub

%>