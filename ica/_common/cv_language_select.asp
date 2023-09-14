
<% 
    objConn.Close
    objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & objExpertDB.DatabasePath & ";"

    Set objTempRs = GetDataRecordsetSP("usp_ExpertGetAllProfilesSelect", Array( _
        Array(, adInteger, , iCvID) _
    ))
%>

<% If Not objTempRs.Eof Then %>
    <div class="box grey gadget">
        <h3>Expert Language</h3>

		<div class="content">		
			<div align="center">
                <select name="cv_language" id="cv_language" style="width: 152px;margin-bottom:5px;" onchange="loadCv();">
                    <% If Not IsNull(objTempRs("id_Expert_Eng")) Then %> <option value="<%= objTempRs("uid_Expert_Eng") %>" <% If sCvUID = objTempRs("uid_Expert_Eng") Then %>selected <% End If %> >Eng</option> <% End If %>
                    <% If Not IsNull(objTempRs("id_Expert_Fra")) Then %> <option value="<%= objTempRs("uid_Expert_Fra") %>" <% If sCvUID = objTempRs("uid_Expert_Fra") Then %>selected <% End If %> >Fra</option> <% End If %>
                    <% If Not IsNull(objTempRs("id_Expert_Spa")) Then %> <option value="<%= objTempRs("uid_Expert_Spa") %>" <% If sCvUID = objTempRs("uid_Expert_Spa") Then %>selected <% End If %> >Spa</option> <% End If %>
                </select>
            </div>
		</div>
		
	</div>    
<% End If %>

<script>
    function loadCv()
    {
        var uid = $('#cv_language option:selected').val();
        if (window.location.pathname.indexOf('cv_verify') > -1) {
            window.location = "cv_verify.asp?uid=" + uid;
        }
        else {
            window.location = "register6.asp?uid=" + uid;
        }
    }
</script>