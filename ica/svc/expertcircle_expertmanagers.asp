<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>


<!--#include file="../dbc.asp"-->
<!--#include file="../fnc.asp"-->
<!--#include file="../fnc_exp.asp"-->
<%
    Dim sFirstName, sLastName, iCompanyID, iExpertCircleID, iCurrentUserID

    iExpertCircleID = Request.QueryString("id")
    iCompanyID = Request.QueryString("companyId")
    iCurrentUserID = CInt(Request.QueryString("userId"))    

    Set objTempRs2 = GetDataRecordsetSP("usp_" & sIcaServerSqlPrefix & "GetICAExpertManagerByCompany", Array( _
        Array(, adInteger, , iCompanyID)))

%>

<script>
    function updateExpertManager(circleid) {        
        var userid = $('#ex-changeexpertmanagers-' + circleid + ' :selected')[0].id;
        var managersName = $('#ex-changeexpertmanagers-' + circleid + ' :selected')[0].text;
        
        $.ajax({
            url: '../../svc/expertcircle_updateexpertmanager.asp',
            data: { iuserid: userid, icircleid: circleid },
            cache: false,
            success: function (data) {      
                onUpdateManagerComplete(
                    data.userid,
                    data.circleid,
                    $('#ex-changeexpertmanagers-' + data.circleid + ' :selected')[0].text
                );
            },
            error: function (jqXHR, textStatus, err) {                
                alert('Error saving Expert Manager: ');
            }
        });
    }

    function cancelUpdate(circleId) {        
        $('#uem-' + circleId).remove();
        $('#updateExpertLink-' + circleId).show();
    }
</script>

<div id="uem-<%=iExpertCircleID%>" >
    <div style="display:inline-block;">
        
            <small><%= objTempRs2("addedByUserFullName")%></small>
        <select id="ex-changeexpertmanagers-<%=iExpertCircleID%>">
            <%  While Not objTempRs2.Eof %>
                    <option id='<%= objTempRs2("IDUSER")%>' <%If CInt(iCurrentUserID) = CInt(objTempRs2("IDUSER")) Then%>selected<%End If%> ><%= objTempRs2("addedByUserFullName")%></option>
            <%      objTempRs2.MoveNext
                WEnd 
            %>
        </select>
    </div>
    <div style="display: inline;">
        <img src="/Resources/Images/cross.png" style="width:16px;height:16px;" onclick="cancelUpdate(<%=iExpertCircleID%>);" />
        <img src="/Resources/Images/green-tick.png" style="width:16px;height:16px;" onclick="updateExpertManager(<%=iExpertCircleID%>);" />
    </div>
</div>
    
<% CloseDBConnection %>