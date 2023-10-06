<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Languages.
'
'--------------------------------------------------------------------
%>
<!--#include virtual="/_template/asp.header.nocache.asp"-->
<!--#include file="_data/datLanguage.asp"-->
<!--#include file="_data/datLngLevel.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()

' Log: 34 - Update expert
If Request.Form()>"" Then
	iLogResult = LogActivity(34, "ExpertID=" & Cstr(iExpertID) & " SavedStep: 5", "", "")
End If

Dim objConnCustom
Set objConnCustom = Server.CreateObject("ADODB.Connection")
objConnCustom.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

' The page is accessible only if expert id is valid
If iExpertID<=0 Then Response.Redirect(sApplicationHomePath & "register.asp" & sParams)

sParams=ReplaceUrlParams(sParams, "lngid")

Dim iExpLngID, iLanguageID, iReadingLevel, iSpeakingLevel, iWritingLevel, iTotalLanguages, j
Dim arrExpNativeLanguages()

iExpLngID=Request.QueryString("lngid")
If IsNumeric(iExpLngID) And iExpLngID>"" Then
	iExpLngID=CLng(iExpLngID)
Else
	iExpLngID=0
End If

If IsNumeric(iExpLngID) And iExpLngID>"" And sAction="delete" Then
	' Deleting data on projects 	
	objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvLanguageOtherDelete", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpLngID)))
	Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "lngid"))
End If

If Request.Form()>"" then
	If Request.Form("newloc")<>"" then
	' Saving national languages
	objTempRs=DeleteRecordSPWithConn(objConnCustom, "usp_ExpCvvLanguageNativeDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , Null)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSPWithConn(objConnCustom, "usp_ExpCvvLanguageNativeInsert", Array( _
		Array(, adVarChar, 2000, CheckString(Request.Form("newloc"))), _
		Array(, adInteger, , iExpertID)), _
		Array( Array(, adInteger)))
	iTotalLanguages=objTempRs(0)
	Set objTempRs=Nothing
	End If

	If Request.Form("exp_lng_name")>0 Then
	iExpLngID=CheckString(Request.Form("id_Lan"))
	iLanguageID=CheckString(Request.Form("exp_lng_name"))
	iReadingLevel=CheckString(Request.Form("exp_lng_read"))
	iSpeakingLevel=CheckString(Request.Form("exp_lng_speak"))
	iWritingLevel=CheckString(Request.Form("exp_lng_write"))

	If IsNumeric(iExpLngID) And iExpLngID>"" And iExpLngID<>"0" Then
		objTempRs=UpdateRecordSPWithConn(objConnCustom, "usp_ExpCvvLanguageOtherUpdate", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , iExpLngID), _
			Array(, adSmallInt, , iLanguageID), _
			Array(, adSmallInt, , iReadingLevel), _
			Array(, adSmallInt, , iSpeakingLevel), _
			Array(, adSmallInt, , iWritingLevel)))
	Else
		objTempRs=InsertRecordSPWithConn(objConnCustom, "usp_ExpCvvLanguageOtherInsert", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adSmallInt, , iLanguageID), _
			Array(, adSmallInt, , iReadingLevel), _
			Array(, adSmallInt, , iSpeakingLevel), _
			Array(, adSmallInt, , iWritingLevel)),"-")
	End If
	End If

	If Request.Form("exp_lng_continue")="0" then
		Response.Redirect "register4.asp" & sParams
	Else
		Response.Redirect "register5.asp" & sParams
        End if   
End If %>
<!--#include virtual="/_template/html.header.start.asp"-->
<script type="text/javascript" src="/_scripts/js/main.js"></script>
<script language="JavaScript">
<!-- 
function validateForm() {
	if (document.locations.newloc.value=="") {
		alert("<% =GetLabel(sCvLanguage, "Please select your native language") %>"); return;
	}
	if (document.locations.exp_lng_name.selectedIndex >0) {
		AddLanguage(1);
	} else {
		document.locations.submit();
	}
}

function AddLanguage(cont_next) { 
	if (document.locations.exp_lng_name.selectedIndex==0) {
		alert("<% =GetLabel(sCvLanguage, "Please choose a language") %>");  return;
	}
	if (document.locations.exp_lng_read.selectedIndex==0 || document.locations.exp_lng_speak.selectedIndex==0 || document.locations.exp_lng_write.selectedIndex==0) {
		alert("<% =GetLabel(sCvLanguage, "Please choose the levels of your knowledge") %>");   return;
	}
	if (cont_next!=1) {
		document.locations.exp_lng_continue.value="0";
	}
	document.locations.submit();
}

//////////////////:starts from here /////////////:::
/////////////////////////////////////
    function moveloc(intDirection)
    {
    	var heading = '';
    	var msg = '';
    	var flag;
    	var arrnew = new Array();
    
    	with (document.locations){
    		if (intDirection){
    			//Add it to mylocations
    			msg = '';
    			for(var x=0;x<locations.length;x++){
    				var opt = locations.options[x];
    				if (opt.selected){
    					flag = 1;
    					//if more then 20 then alert and exit.
    					if (mylocations.length > 19){
    						alert("You are only allowed 20 languages.");
    						break;	
    					}
    					//check if option exists if not add it
    					for (var y=0;y<mylocations.length;y++){
    						var myopt = mylocations.options[y];
    						if (myopt.value == opt.value){	
    							flag = 0;
    						}
    					}
    					if (flag){
    						//This is not a duplicate so add it to the select box.
    						mylocations.options[mylocations.options.length] = new Option(opt.text, opt.value, 0, 0);
    					}
    				}
    			}
    		}else{
    			//Delete it from my locations
    			for(var x=mylocations.length-1;x>=0;x--){
    				var opt = mylocations.options[x];
    				if (opt.selected){									
    					//Remove it from the select box
    					mylocations.options[x] = null;
    				}
    			}		
    		}
    
    		//Fill hidden field with new values
    		for (var y=0;y<mylocations.length;y++){
    			arrnew[y] = mylocations.options[y].value
    		}			
    		newloc.value = arrnew.join();
    
    	}
    }
    
    function InitLoc(){
    	var arrnew = new Array();
    	with (document.locations){
    		//Clear first line of mylocations (Netscape)
    		if (mylocations.options[0].value == 0){
    			mylocations.options[0] = null;	
    		}
    		
    		//Fill hidden field with new values
    		for (var y=0;y<mylocations.length;y++){
    			arrnew[y] = mylocations.options[y].value
    		}			
    		newloc.value = arrnew.join();
    
    	}
    	return 1;
    }
////////////////:
// stop hiding -->
</script>
</head>
<body>
<div id="holder">
	<!-- header -->
	<!--#include virtual="/_template/page.header.asp"-->

	<!-- content -->
	<div id="content" class="searchform">
	<% 
	If Not bIsMyCV Then 
		%><h2 class="service_title">Curriculum Vitae. <span class="service_slogan">Expert ID: <% =objExpertDB.DatabaseCode %><%=iExpertID%></span></h2><br/>
		<% 
	End If

	ShowRegistrationProgressBar "CV", 5
	%>
		
	<form method="post" name="locations">
	<input type="hidden" name="exp_lng_continue" value="1">
	<input type="hidden" name="newloc" value="">
	<% Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adVarChar, 10, "native")))
	Redim arrExpNativeLanguages(0)
	i=1
	While Not objTempRs.Eof 
		ReDim Preserve arrExpNativeLanguages(i)
		arrExpNativeLanguages(i)=objTempRs("id_Language")
		objTempRs.MoveNext
		i=i+1
	WEnd
	objTempRs.Close	%>

  <!-- Native Languages -->
  
<% ShowMessageStart "info", 580 %>
<% =GetLabel(sCvLanguage, "Add selected language") %>
<% ShowMessageEnd %><br/>

		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Native languages") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td class="field splitter"><label for="exp_otr_name"><% =GetLabel(sCvLanguage, "Languages") %></label></td>
		<td class="value blue">
			<table><tr>
			<td>
			<select name="locations" size="7" multiple style="width:145px;">
			<% For j=1 to Ubound(arrLanguageID) %>
			<option value="<%=arrLanguageID(j)%>"<% For i=1 to UBound(arrExpNativeLanguages)%><% if arrExpNativeLanguages(i)=arrLanguageID(j) then %> selected<%end if%><%next%>><%=arrLanguageTitle(j)%></option>
			<% Next %>
			</select>
			</td>
			<td align="center">
			<img src="/image/x.gif" width="70" height="1"><br/>
			<a href="javascript:moveloc(1)"><% =GetLabel(sCvLanguage, "Add") %></a><br/><br/><br><a href="javascript:moveloc(0)"><% =GetLabel(sCvLanguage, "Remove") %></a></p>
			</td>
			<td>
			<select name="mylocations" size="7" multiple style="width:145px">
			<option value="" selected>---------------------------------</option>
			<% For j=1 to Ubound(arrLanguageID) %>
			<% For i=1 to UBound(arrExpNativeLanguages)%><%	If arrExpNativeLanguages(i)=arrLanguageID(j) then %>
			<option value="<%=arrLanguageID(j)%>"><%=arrLanguageTitle(j)%></option>
			<% End If %><% Next %>
			<% Next %>
			</select>
			</td>
			<td>&nbsp;&nbsp;<span class="fcmp">*</span></td>
			</tr></table>
		</td>
		</tr>
		</table>
		</div>
		
<% ShowMessageStart "info", 580 %>
<% =GetLabel(sCvLanguage, "Choose a language and specify your level...") %>
<% ShowMessageEnd %><br/>

		<div class="box search blue">
		<h3><span class="left">&nbsp;</span><span class="right">&nbsp;</span><% =GetLabel(sCvLanguage, "Foreign languages") %></h3>
		<table class="search_form" width="100%" cellspacing=0 cellpadding=0 border=0>
	<% ' List of languages
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adVarChar, 10, "other")))
	If Not objTempRs.eof Then %>
	<tr>
	<td colspan=2>
		<table class="results" style="border-left:0; border-right:0;">
		<tr class="tr_results">
		<th class="number"><p>N.</p></td>
		<th width=240><p><% =GetLabel(sCvLanguage, "Language") %></p></td>
		<th width=100><p><% =GetLabel(sCvLanguage, "Reading") %></p></td>
		<th width=100><p><% =GetLabel(sCvLanguage, "Speaking") %></p></td>
		<th width=100><p><% =GetLabel(sCvLanguage, "Writing") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Modify") %></p></td>
		<th width=15><p><% =GetLabel(sCvLanguage, "Delete") %></p></td>
		</tr>	
		<% i=1
		 while not objTempRs.EOF %>
		<tr class="tr_results<% If i Mod 2 = 0 Then %> odd<% End If %>">
		<td><p align="center"><%=i%>.</td>
		<td><p><a href="<%=sScriptFileName & AddUrlParams(sParams, "lngid=" & objTempRs("id_ExpLan"))%>&act=update"><% If objTempRs("id_ExpLan")=iExpLngID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %>
		    	<% 
		    	If sCvLanguage = cLanguageFrench Then
		    		Response.Write objTempRs("lngNameFra")
		    	ElseIf sCvLanguage = cLanguageSpanish Then
		    		Response.Write objTempRs("lngNameSpa")
		    	Else
		    		Response.Write objTempRs("lngNameEng")
		    	End If 
		    	%>
		</a></td>
		<td><p><%=arrLanguageLevelTitle(objTempRs("exlReading"))%>&nbsp;</td>
		<td><p><%=arrLanguageLevelTitle(objTempRs("exlSpeaking"))%>&nbsp;</td>
		<td><p><%=arrLanguageLevelTitle(objTempRs("exlWriting"))%>&nbsp;</td>
		<td align="center"><% If objTempRs("id_ExpLan")=iExpLngID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="<%=sScriptFileName & AddUrlParams(sParams, "lngid=" & objTempRs("id_ExpLan"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		<td align="center"><a href="<%=sScriptFileName & AddUrlParams(sParams, "lngid=" & objTempRs("id_ExpLan"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		</tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd %>
		</table>
	</td>
	</tr>
	<% End If 
	objTempRs.Close %>
    
<% If iExpLngID>0 then
	Set objTempRs=GetDataRecordsetSPWithConn(objConnCustom, "usp_ExpCvvLanguageInfoSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpLngID)))
	If Not objTempRs.Eof Then
	iLanguageID=objTempRs("id_Language")
	iReadingLevel=objTempRs("exlReading")
	iSpeakingLevel=objTempRs("exlSpeaking")
	iWritingLevel=objTempRs("exlWriting")
	End If
	objTempRs.Close
End If %>
		<tr>
		<td class="field splitter"><label for="exp_lng_name"><% =GetLabel(sCvLanguage, "Language") %></label></td>
		<td class="value blue"><select id="exp_lng_name" name="exp_lng_name" size="1" style="width:200px;">
		<option value="0" selected> <% =GetLabel(sCvLanguage, "Please select") %> </option>
		<%		
		For j=1 to Ubound(arrLanguageID)
		If arrLanguageID(j)=iLanguageID then
			Response.Write("<option value=" & arrLanguageID(j) & " selected>"& arrLanguageTitle(j) & "</option>")
		Else
			Response.Write("<option value=" & arrLanguageID(j) & ">"& arrLanguageTitle(j) & "</option>")
		End if 
		Next %>
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_lng_read"><% =GetLabel(sCvLanguage, "Reading") %></label></td>
		<td class="value blue"><select id="exp_lng_read" name="exp_lng_read" size="1" style="width:200px;">
		<option value="0" selected> <% =GetLabel(sCvLanguage, "Please select") %> </option>
		<% For j=1 to Ubound(arrLanguageLevelID) 
	        If arrLanguageLevelID(j)=iReadingLevel then 
			Response.Write("<option value=" & arrLanguageLevelID(j) & " selected>"& arrLanguageLevelTitle(j) &"</option>")
		Else
			Response.Write("<option value=" & arrLanguageLevelID(j) & ">"& arrLanguageLevelTitle(j) &"</option>")
		End If
		Next %>
		</select></td>
		</tr>
		<tr>
		<td class="field splitter"><label for="exp_lng_speak"><% =GetLabel(sCvLanguage, "Speaking") %></label></td>
		<td class="value blue"><select id="exp_lng_speak" name="exp_lng_speak" size="1" style="width:200px;">
		<option value="0" selected> <% =GetLabel(sCvLanguage, "Please select") %> </option>
		<% For j=1 to Ubound(arrLanguageLevelID) 
	        If arrLanguageLevelID(j)=iSpeakingLevel then 
			Response.Write("<option value=" & arrLanguageLevelID(j) & " selected>"& arrLanguageLevelTitle(j) &"</option>")
		Else
			Response.Write("<option value=" & arrLanguageLevelID(j) & ">"& arrLanguageLevelTitle(j) &"</option>")
		End If
		Next %>
		</select></td>
		</tr>
		<tr class="last">
		<td class="field splitter"><label for="exp_otr_name"><% =GetLabel(sCvLanguage, "Writing") %></label></td>
		<td class="value blue"><select name="exp_lng_write" size="1" style="width:200px;">
		<option value="0" selected> <% =GetLabel(sCvLanguage, "Please select") %> </option>
		<% For j=1 to Ubound(arrLanguageLevelID) 
	        If arrLanguageLevelID(j)=iWritingLevel then 
			Response.Write("<option value=" & arrLanguageLevelID(j) & " selected>"& arrLanguageLevelTitle(j) &"</option>")
		Else
			Response.Write("<option value=" & arrLanguageLevelID(j) & ">"& arrLanguageLevelTitle(j) &"</option>")
		End If
		Next %></select></td>
		</tr>
		</table>
		</div>
		
		<div class="spacebottom">
		<a href="javascript:AddLanguage(0);"><img class="button first" src="<% =sHomePath %>image/bte_<% if iExpLngID>0 then %>save<% Else %>add<% End If %>lang.gif" name="Add this project to the list of managed projects" alt="Append this record about language knowledge" border=0></a>
		<a href="javascript:validateForm();"><img class="button last" src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save and continue"  border=0></a>
		</div>
		<input type="hidden" name="id_Lan" value="<%=iExpLngID%>">
		</form>

	</div>
</div>
	<!-- footer -->
	<!--#include file="../_template/page.footer.asp"-->
<% CloseDBConnection %>
<script>
    <!--
    InitLoc();
    //-->
</script>
</body>
<!--#include file="../_template/html.footer.asp"-->
