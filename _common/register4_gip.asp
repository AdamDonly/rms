<% 
'--------------------------------------------------------------------
'
' CV registration.
' Full format. Languages.
'
'--------------------------------------------------------------------
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.CacheControl = "no-cache" 
%>
<!--#include file="_data/datLanguage.asp"-->
<!--#include file="_data/datLngLevel.asp"-->
<%
If sApplicationName<>"expert" Then
	CheckUserLogin sScriptFullNameAsParams
End If
CheckExpertID()
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
	objTempRs=UpdateRecordSP("usp_ExpCvvLanguageOtherDelete", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , iExpLngID)))
	Response.Redirect(sScriptFileName & ReplaceUrlParams(sParams, "lngid"))
End If

If Request.Form()>"" then
	If Request.Form("newloc")<>"" then
	' Saving national languages
	objTempRs=DeleteRecordSP("usp_ExpCvvLanguageNativeDelete", Array( _ 
		Array(, adInteger, , iExpertID), _
		Array(, adInteger, , Null)))
	Set objTempRs=Nothing
	objTempRs=GetDataOutParamsSP("usp_ExpCvvLanguageNativeInsert", Array( _
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
		objTempRs=UpdateRecordSP("usp_ExpCvvLanguageOtherUpdate", Array( _
			Array(, adInteger, , iExpertID), _
			Array(, adInteger, , iExpLngID), _
			Array(, adSmallInt, , iLanguageID), _
			Array(, adSmallInt, , iReadingLevel), _
			Array(, adSmallInt, , iSpeakingLevel), _
			Array(, adSmallInt, , iWritingLevel)))
	Else
		objTempRs=InsertRecordSP("usp_ExpCvvLanguageOtherInsert", Array( _
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

<html>
<head>
<title><% =GetLabel(sCvLanguage, "CV registration") %>. <% =GetLabel(sCvLanguage, "Languages") %></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" type="text/css" href="<% =sHomePath %>styles.css">
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

<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0  marginheight=0 marginwidth=0>
<% ShowTopMenu %>
<% ShowRegistrationProgressBar "CV", 5 %>
<br/>
  <!-- Other section -->
	<form method="post" name="locations">
	<input type="hidden" name="exp_lng_continue" value="1">
	<input type="hidden" name="newloc" value="">

	<% Set objTempRs=GetDataRecordsetSP("usp_ExpCvvLanguageSelect", Array( _
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
<% ShowMessageEnd %>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br /><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "NATIVE LANGUAGES") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="#FFE2E2"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170 valign="top"><p class="ftxt"><% =GetLabel(sCvLanguage, "Languages") %><br /><img src="<% =sHomePath %>image/x.gif" width=140 height=1><br /></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=170 valign="top"><img src="x.gif" width=1 height=5><br />
		&nbsp;&nbsp;<select name="locations" size="7" multiple style="width:155px;">
		<% For j=1 to Ubound(arrLanguageID) %>
			<option value="<%=arrLanguageID(j)%>"<% For i=1 to UBound(arrExpNativeLanguages)%><% if arrExpNativeLanguages(i)=arrLanguageID(j) then %> selected<%end if%><%next%>><%=arrLanguageTitle(j)%></option>
		<% Next %>
		</select>
		<img src="<% =sHomePath %>image/x.gif" width=170 height=1 vspace=3><br />
		</td>

		<td bgcolor="<%=colFormBodyText%>" width=32 valign="center" align="center"><p>
		<a href="javascript:moveloc(1)"><% =GetLabel(sCvLanguage, "Add") %></a><br /><br /><br /><a href="javascript:moveloc(0)"><% =GetLabel(sCvLanguage, "Remove") %></a></p>
		</td>

		<td bgcolor="<%=colFormBodyText%>" width=170 valign="top"><img src="x.gif" width=1 height=5><br />
		<select name="mylocations" size="7" multiple style="width:155px">
                   <OPTION value="" SELECTED>---------------------------------</OPTION>
   		<% For j=1 to Ubound(arrLanguageID) %>
   		<% For i=1 to UBound(arrExpNativeLanguages)%><%	If arrExpNativeLanguages(i)=arrLanguageID(j) then %>
		 <option value="<%=arrLanguageID(j)%>"><%=arrLanguageTitle(j)%></option>
		 <%end if%><%next%>
		<% Next %>
                </select>
		</td>
		<td bgcolor="<%=colFormBodyText%>" width=39 valign="center"><span class="fcmp">*</span>&nbsp;&nbsp;</td>
		</tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br />

<% ShowMessageStart "info", 580 %>
<% =GetLabel(sCvLanguage, "Choose a language and specify your level...") %>
<% ShowMessageEnd %>

  <!-- Foreign Languages -->
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br /><p class="fttl"><img src="<% =sHomePath %>image/<% =imgFormBullet %>" width=7 height=7 align=left vspace=3 hspace=8><% =GetLabel(sCvLanguage, "FOREIGN LANGUAGES") %></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<% ' List of languages
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvLanguageSelect", Array( _
		Array(, adInteger, , iExpertID), _
		Array(, adVarChar, 10, "other")))
	If Not objTempRs.eof Then %>
	<tr>
	<td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=579 colspan=2 bgcolor="<%=colFormBodyRight%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr>
		<td>
		<table cellspacing=1 cellpadding=4 align="center" width="100%" border=0 bgcolor="<%=colFormBodyRight%>">
		<tr height=20>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p>N.</td>
			<td width=240 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Language") %></td>
			<td width=100 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Reading") %></td>
			<td width=100 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Speaking") %></td>
			<td width=100 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Writing") %></td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Modify") %></td>
			<td width=15 bgcolor="<%=colFormHeaderTop%>"><p><% =GetLabel(sCvLanguage, "Delete") %></td>
		</tr>	
		<% i=1
		 while not objTempRs.EOF %>
		  <tr height=20>
		    <td bgcolor="<%=colFormBodyText%>"><p ><%=i%></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><a href="<%=sScriptFileName & AddUrlParams(sParams, "lngid=" & objTempRs("id_ExpLan"))%>&act=update"><% If objTempRs("id_ExpLan")=iExpLngID Then %><b><img src="<% =sHomePath %>image/vn_v.gif" width=8 height=12 border=0 hspace=0 align="left"><% End If %>
		    	<% 
		    	If sCvLanguage = cLanguageFrench Then
		    		Response.Write objTempRs("lngNameFra")
		    	Else
		    		Response.Write objTempRs("lngNameEng")
		    	End If 
		    	%>
		    	</a></td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=arrLanguageLevelTitle(objTempRs("exlReading"))%>&nbsp;</td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=arrLanguageLevelTitle(objTempRs("exlSpeaking"))%>&nbsp;</td>
		    <td bgcolor="<%=colFormBodyText%>"><p><%=arrLanguageLevelTitle(objTempRs("exlWriting"))%>&nbsp;</td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><% If objTempRs("id_ExpLan")=iExpLngID Then %><img src="<% =sHomePath %>image/vn_updte.gif" width=15 height=15 border=0 hspace=0 alt="Updating" align="center"><% Else %><a href="<%=sScriptFileName & AddUrlParams(sParams, "lngid=" & objTempRs("id_ExpLan"))%>&act=update"><img src="<% =sHomePath %>image/vn_updt.gif" width=15 height=15 border=0 hspace=0 alt="Update this record" align="center"></a><% End If %></td>
		    <td bgcolor="<%=colFormBodyText%>" align="center"><a href="<%=sScriptFileName & AddUrlParams(sParams, "lngid=" & objTempRs("id_ExpLan"))%>&act=delete"><img src="<% =sHomePath %>image/vn_del.gif" width=15 height=15 border=0 hspace=0 alt="Delete this record" align="center"></a></td>
		  </tr>
		<% i=i+1
		objTempRs.MoveNext
		WEnd %>
		</table>
		</td>
		</tr>
		</table>
	</td>
	</tr>

	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<% End If 
	objTempRs.Close %>
    
<% If iExpLngID>0 then
	Set objTempRs=GetDataRecordsetSP("usp_ExpCvvLanguageInfoSelect", Array( _
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

	<a name="A">
	<tr><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table width="100%" cellspacing=0 cellpadding=0 border=0>
		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Language") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407><img src="x.gif" width=1 height=6><br />
		&nbsp;&nbsp;<select name="exp_lng_name" size="1" style="width:200px;">
		<option value="0" selected><% =GetLabel(sCvLanguage, "Please select") %></option>
		<%		
		For j=1 to Ubound(arrLanguageID)
		If arrLanguageID(j)=iLanguageID then
			Response.Write("<option value=" & arrLanguageID(j) & " selected>"& arrLanguageTitle(j) & "</option>")
		Else
			Response.Write("<option value=" & arrLanguageID(j) & ">"& arrLanguageTitle(j) & "</option>")
		End if 
		Next %>
		</select></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Reading") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_lng_read" size="1" style="width:200px;">
		<option value="0" selected><% =GetLabel(sCvLanguage, "Please select") %></option>
		<% For j=1 to Ubound(arrLanguageLevelID) 
	        If arrLanguageLevelID(j)=iReadingLevel then 
			Response.Write("<option value=" & arrLanguageLevelID(j) & " selected>"& arrLanguageLevelTitle(j) &"</option>")
		Else
			Response.Write("<option value=" & arrLanguageLevelID(j) & ">"& arrLanguageLevelTitle(j) &"</option>")
		End If
		Next %>
		</select></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Speaking") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_lng_speak" size="1" style="width:200px;">
		<option value="0" selected><% =GetLabel(sCvLanguage, "Please select") %></option>
		<% For j=1 to Ubound(arrLanguageLevelID) 
	        If arrLanguageLevelID(j)=iSpeakingLevel then 
			Response.Write("<option value=" & arrLanguageLevelID(j) & " selected>"& arrLanguageLevelTitle(j) &"</option>")
		Else
			Response.Write("<option value=" & arrLanguageLevelID(j) & ">"& arrLanguageLevelTitle(j) &"</option>")
		End If
		Next %>
		</select></td></tr>

		<tr><td width=170><p class="ftxt"><% =GetLabel(sCvLanguage, "Writing") %></td>
		<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=24></td>
		<td bgcolor="<%=colFormBodyText%>" width=407>
		&nbsp;&nbsp;<select name="exp_lng_write" size="1" style="width:200px;">
		<option value="0" selected><% =GetLabel(sCvLanguage, "Please select") %></option>
		<% For j=1 to Ubound(arrLanguageLevelID) 
	        If arrLanguageLevelID(j)=iWritingLevel then 
			Response.Write("<option value=" & arrLanguageLevelID(j) & " selected>"& arrLanguageLevelTitle(j) &"</option>")
		Else
			Response.Write("<option value=" & arrLanguageLevelID(j) & ">"& arrLanguageLevelTitle(j) &"</option>")
		End If
		Next %></select>
		<img src="<% =sHomePath %>image/x.gif" width=358 height=1 vspace=2><br />
		</td></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyText%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<% =colFormHeaderSplitter %>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table><br />

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=380 align=left>
	<img src="<% =sHomePath %>image/x.gif" width=170 height=1><a href="javascript:AddLanguage(0);"><img src="<% =sHomePath %>image/bte_<% if iExpLngID>0 then %>save<% Else %>add<% End If %>lang.gif" name="Add this project to the list of managed projects" alt="Append this record about language knowledge" border=0></a>
	<input type="hidden" name="id_Lan" value="<%=iExpLngID%>">
	</td>
	<td width=200 height=1 align="right"><a href="javascript:validateForm();"><img src="<% =sHomePath %>image/bte_savecont.gif" name="Continue" alt="Save and continue"  border=0></a></td>
	</tr>
	</form>
	</table><br />

<SCRIPT LANGUAGE=javascript>
    <!--
    InitLoc();
    //-->
</SCRIPT>

<% CloseDBConnection %>
</body>
</html>
