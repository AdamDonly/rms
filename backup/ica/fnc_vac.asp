<%
Sub RenderVacanciesSelector(expertId, expertUid, databaseId)
	' get all "my vacancies":
	Dim bHasLinkToVacancy, sLinkedPositionId
	bHasLinkToVacancy = 0
	sLinkedPositionId = ""
	databaseId = Replace(databaseId, "-", "")
	%><div class="vacancy-links-container box grey gadget">
		<h3>Linked to vacancies:</h3>
		<%
		Set objTempRs = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "ExpertVacanciesLinksSelect", Array( _
			Array(, adVarChar, 40, expertUid), _
			Array(, adInteger, , iUserID)))
		If Not objTempRs.Eof Then
			While Not objTempRs.Eof
				sLinkedPositionId = CStr(objTempRs("IDPOSITION"))
				bHasLinkToVacancy = 1
				%><div class="vacancy-link-holder" data-positionId="<%=sLinkedPositionId %>">
					<a href="http://<% =sIcaServer %>/Public/ICAVacancyDetails?id=<%=sLinkedPositionId %>" target="_blank" class="vacancy-title"><%=objTempRs("POSITION_TITLE") %></a>&nbsp;
					<a href="javascript:void(0)" class="btn-remove-expert-to-position grey-button floatRight" data-expertId="<%=expertId %>" data-expertUid="<%=expertUid %>" data-positionId="<%=sLinkedPositionId %>" data-cvdatabase="<%=databaseId %>">Remove Link</a>
				</div>
				<%
				objTempRs.MoveNext
			Wend
		Else
			%><p id="no-linked-positions">Currently not linked to any position.</p><%
		End If
		objTempRs.Close
		Set objTempRs = Nothing
		' selector for new links:
		%><div class="vacancy-link-holder" data-positionId="0">
			<div>Link this expert to a vacancy:</div>
			<select name="vacanciesSelector" data-positionId="0" data-expertId="<%=expertId %>" data-expertUid="<%=expertUid %>" data-cvdatabase="<%=databaseId %>">
				<option value=""></option>
				<%=BuildVacanciesSelectorOptions(0) %>
			</select>
		</div>
		<br class="clear" />
		<a class="vacancy-link-closer grey-button floatRight" href="javascript:void(0)">close</a>
		<br class="clear" />
		<br/>
	</div>
	<a href="javascript:void(0)" class="vacancy-link-opener">Link<% If bHasLinkToVacancy = 1 Then %>ed<% End If %> to Vacancy</a>
	<style type="text/css">
	.vacancy-link-opener {display:block;float:right;padding:2px 5px;border-radius:3px;border:1px solid #999;margin:0 5px;text-decoration:none;}
	.vacancy-link-closer {float:right;margin-right:5px;}
	.vacancy-link-opener:hover {text-decoration:none;}
	.vacancy-links-container {display:none;width:50% !important;position:absolute;top:100px;left:25%;box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19);}
	.vacancy-links-container .link-to-txt {display:block;font-size:1.1em;font-weight:bold;padding:5px 0;}
	.vacancy-link-holder {clear:both;padding:5px;border-bottom:1px solid #999;margin-bottom:10px;}
	.vacancy-link-holder select[name="vacanciesSelector"] {width:100%;height:25px;}
	.vacancy-link-holder .vacancy-title {line-height:20px;}
	</style>
	<script type="text/javascript">
	$(function () {
		$('.vacancy-link-opener').click(function () {
			if ($('.vacancy-links-container').is(':visible'))
			{
				$('.vacancy-links-container').hide();
			}
			else 
			{
				$('.vacancy-links-container').show();

				$('select[name="vacanciesSelector"]').change(function () {
					if ($(this).val() != '')
					{
						linkExpertToVacancy($(this).val(), $(this).attr('data-expertId'), $(this).attr('data-expertUid'), $(this).attr('data-type'), $(this).attr('data-cvdatabase'), $(this).find('option:selected').text());
					}
				});
				initRemoveLinkButton('');

				$('.vacancy-link-closer').click(function () {
					$('.vacancy-links-container').hide();
				});
			}
		});
	});

	function initRemoveLinkButton(positionId)
	{
		$('.btn-remove-expert-to-position[data-positionId' + (positionId != "" ? ('="' + positionId + '"') : '') + ']').click(function (e) {
			e.preventDefault();
			e.stopPropagation();
			if (confirm("Are you sure you want to remove the link between the expert and the vacancy?"))
			{
				deleteExpertToVacancyLink($(this).attr('data-positionId'), $(this).attr('data-expertId'), $(this).attr('data-expertUid'));
			}
		});
	}

	function linkExpertToVacancy(positionId, expertId, expertUid, linkValue, cvDatabaseId, positionTitle) {
		$.ajax({
			cache: false,
			url: '/svc/positionexpertlinkupdate.asp?positionid=' + positionId + '&expertId=' + expertId + '&expertUid=' + expertUid + '&linkValue=' + linkValue + '&expertDatabaseId=' + cvDatabaseId,
			expertUid: expertUid,
			expertId: expertId,
			positionId: positionId,
			cvDatabaseId: cvDatabaseId,
			positionTitle: positionTitle,
			linkValue: linkValue,
			success: function (result) {
				if ((typeof result) != 'undefined' && result != null && result.indexOf('OK') > -1) {
					if ($('.vacancy-link-holder[data-positionId="' + this.positionId + '"]').length < 1)
					{
						$('<div class="vacancy-link-holder" data-positionId="' + this.positionId + '">' +
							'<a href="http://<% =sIcaServer %>/Public/ICAVacancyDetails?id=' + this.positionId + '" target="_blank" class="vacancy-title">' + this.positionTitle + '</a>&nbsp;' +
							'<a href="javascript:void(0)" class="btn-remove-expert-to-position grey-button floatRight" data-expertId="' + this.expertId + '" data-expertUid="' + this.expertUid + '" data-positionId="' + this.positionId + '" data-cvdatabase="' + this.cvDatabaseId + '">Remove Link</a>' +
						'</div>').insertBefore($('.vacancy-link-holder[data-positionId="0"]'));

						initRemoveLinkButton(this.positionId);
					}
					$('select[name="vacanciesSelector"]').val('');

					$('#no-linked-positions').hide();
				}
				else {
					alert('Error setting link.');
				}
			},
			error: function (jqXHR, textStatus, err) {
				alert('Error setting link.');
			},
			complete: function () {
				
			}
		});
	}

	function setLinkedExpertStatus(positionId, expertId, expertUid, statusValue) {
		$.ajax({
			cache: false,
			url: '/svc/positionexpertstatusupdate.asp?positionid=' + positionId + '&expertId=' + expertId + '&expertUid=' + expertUid + '&statusValue=' + statusValue,
			success: function (result) {
				if (result.indexOf('OK') > -1) {

				}
				else {
					alert('Error setting status.');
				}
			},
			error: function (jqXHR, textStatus, err) {
				alert('Error setting status.');
			},
			complete: function () {
				
			}
		});
	}

	function deleteExpertToVacancyLink(positionId, expertId, expertUid) {
		$.ajax({
			cache: false,
			url: '/svc/positionexpertlinkdelete.asp?positionid=' + positionId + '&expertId=' + expertId + '&expertUid=' + expertUid,
			positionId: positionId,
			success: function (result) {
				if (result.indexOf('OK') > -1) {
					$('.vacancy-link-holder[data-positionId="' + this.positionId + '"]').remove();

					if ($('.vacancy-link-holder[data-positionId]').length < 2) 
					{
						$('#no-linked-positions').show();
					}
				}
				else {
					alert('Error deleting link.');
				}
			},
			error: function (jqXHR, textStatus, err) {
				alert('Error deleting link.');
			},
			complete: function () {
				
			}
		});
	}
	</script>
	<%
End Sub

Function BuildVacanciesSelectorOptions(sSelectedValue)
	Dim sSelectorHtml, iCurrentProjectId
	sSelectorHtml = ""
	iCurrentProjectId = 0
	Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "MyVacanciesSelect", Array( _
		Array(, adInteger, , iUserID)))
	While Not objTempRs2.Eof 
		If iCurrentProjectId <> CDbl(objTempRs2("IDPROJECT")) Then
			iCurrentProjectId = CDbl(objTempRs2("IDPROJECT"))
			If sSelectorHtml <> "" Then
				sSelectorHtml = sSelectorHtml & "</optgroup>"
			End If
			sSelectorHtml = sSelectorHtml & "<optgroup label=""" & objTempRs2("PROJECT_TITLE") & """>"
		End If
		sSelectorHtml = sSelectorHtml & "<option value=""" & objTempRs2("IDPOSITION") & """ "
		If sSelectedValue = CStr(objTempRs2("IDPOSITION")) Then 
			sSelectorHtml = sSelectorHtml & "selected=""selected""" 
		End If 
		sSelectorHtml = sSelectorHtml & ">" & objTempRs2("POSITION_TITLE") & "</option>"
		objTempRs2.MoveNext
	Wend
	If sSelectorHtml <> "" Then
		sSelectorHtml = sSelectorHtml & "</optgroup>"
	End If
	objTempRs2.Close
	Set objTempRs2 = Nothing
	BuildVacanciesSelectorOptions = sSelectorHtml
End Function
%>
