<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>


<!--#include virtual="dbc.asp"-->
<!--#include virtual="fnc.asp"-->

<% 
    Dim sDatabase, sDatabaseId
    sDatabase = Request.QueryString("database")
    sDatabaseId = Request.QueryString("databaseId")

    objConn.Close
    objConn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Data Source=MAGNESA;Initial Catalog=" & sDatabase & ";"

    Dim iEid, sLanguage
    iEid = Request.QueryString("expertId")
    sLanguage = Request.QueryString("language")

    Set objTempRs = GetDataRecordsetSP("usp_ExpertGetAllProfilesSelect", Array( _
        Array(, adInteger, , iEid) _
    ))
%>

<div class="allProfiles">
    <input type="hidden" id="startingLanguage" value="<%= sLanguage %>" />
    <% If objTempRs.RecordCount > 0 Then %>
        <% While Not objTempRs.Eof %>
            <div class="row allProfiles__item">
                currently linked:
                <div class="col-12">
                    <div class="row" data-lang="Eng">
                        <div class="col-4" style="text-align: right;">
                            English 
                        </div>
                        
                        <% If Not IsNull(objTempRs("id_Expert_Eng")) Then %> 
                            <div class="col-4" style="text-align: right;">
                                (<%= objTempRs("id_Database") & "-" & objTempRs("id_Expert_Eng") %>) 
                            </div>
                            <div class="col-4">
                                <a href="#" onclick="removeLink(<%= objTempRs("id_expertLanguage") %>, 'Eng');">remove</a>
                            </div>
                        <% Else %>
                            <div class="col-4" style="text-align: right;">
                                <a href="#" class="updateProfileLink" onclick="addNewLink('Eng');">add link</a>
                            </div>
                            <div class="col-4"></div>
                        <% End If %>
                    </div>
                    
                    <div class="row" data-lang="Fra">
                        <div class="col-4" style="text-align: right;">
                            French
                        </div>
                        <% If Not IsNull(objTempRs("id_Expert_Fra")) Then %> 
                            <div class="col-4" style="text-align: right;">
                                (<%= objTempRs("id_Database") & "-" & objTempRs("id_Expert_Fra") %>) 
                            </div>
                            <div class="col-4">
                                <a href="#" onclick="removeLink(<%= objTempRs("id_expertLanguage") %>, 'Fra');">remove</a>
                            </div>
                        <% Else %> 
                            <div class="col-4" style="text-align: right;">                            
                                <a href="#" class="updateProfileLink" onclick="addNewLink('Fra');">add link</a>
                            </div>
                            <div class="col-4"></div>
                        <% End If %>
                    </div>
                    
                    <div class="row" data-lang="Spa">
                        <div class="col-4" style="text-align: right;">
                            Spanish
                        </div>
                        <% If Not IsNull(objTempRs("id_Expert_Spa")) Then %> 
                            <div class="col-4" style="text-align: right;">
                                (<%= objTempRs("id_Database") & "-" & objTempRs("id_Expert_Spa") %>) 
                            </div>
                            <div class="col-4">
                                <a href="#" onclick="removeLink(<%= objTempRs("id_expertLanguage") %>, 'Spa');">remove</a>
                            </div>
                        <% Else %> 
                            <div class="col-4" style="text-align: right;">                            
                                <a href="#" class="updateProfileLink" onclick="addNewLink('Spa');">add link</a>
                            </div>
                            <div class="col-4"></div>
                        <% End If %>
                    </div>
                </div>
            </div>
            <% objTempRs.MoveNext %> 
        <% WEnd %>
    <% Else %>
            <div style="text-align: center;">No linked profiles</div>
            <div class="col-12" style="margin-top: 5px;">
                <input type="button" class="linkbuttons" onclick="linkExistingOrInitialProfile(1);" value="Create" />
                <input type="button" class="linkbuttons" onclick="$('#modalContainer .close').click();" value="Cancel" />
            </div>
    <% End If %>

    <!-- <div class="row buttons" style="display: none; margin-top: 20px;margin-bottom: 10px;">
        <div class="col-12" style="text-align: center;">
            <div class="btn" data-id="new">new profile</div>
            <div class="btn" data-id="existing">link existing</div>
        </div>
    </div> -->
    <div class="row existingContainer" style="display: none;margin-top:20px;">
        <input type="hidden" id="languageToCreate" />
        <div class="col-6">
            expert id: 
        </div>
        <div class="col-6">
            <input type="text" id="existingId" style="background: #f6f6f6;border: 1px solid #ccc;" />
            <div id="eidvalidation" style="font-size: 12px; color: red;"></div>
        </div>
        <div class="col-12" style="margin-top: 10px;">
            <input type="button" class="linkbuttons" onclick="linkExistingOrInitialProfile(0);" value="Create" />
            <input type="button" class="linkbuttons" onclick="cancelLinkUpdate();" value="Cancel" />
        </div>
    </div>
</div>

<script>

    var database = '<%= sDatabase %>';

    $("#popup").on("hidden.bs.modal", function () {
        $('#popup .modal-body .container').html('');
        $('#popup .modal-body .loaderContainer').html('');
    });
    function cancelLinkUpdate() {
        $('.existingContainer').hide();
        $('#existingId').val();
        $('.row[data-lang]').removeClass('selectedLang');
    }
    function addNewLink(lang) {

        $('.row[data-lang]').removeClass('selectedLang');
        $('.row[data-lang=' + lang + ']').addClass('selectedLang');
        
        $('.buttons').show();
        $('#languageToCreate').val(lang);
        $('.existingContainer').show();
        $('.btn').removeClass('btn--selected');
    }

    function linkExistingOrInitialProfile(initialCreate) {
        profileCreation(initialCreate, function (data) {
            $('#modalContainer .close').click();
            $('#search').click();
        })
    }

    function removeLink(expertLanguageId, language) 
    {
        $.ajax({
            url: "remove-expert-link.asp",
            method: 'POST', 
            data: { expertLanguageId: expertLanguageId, language: language, database: database },
            success: function () { window.location.reload(); },
            error: function () {}
        });
    }

    function checkIdAndLanguage(expertId, language, database, callback) {
        $.ajax({
            url: 'check-id-language.asp',
            method: 'POST', 
            data: { expertId: expertId, language: language, database: database },
            success: function (result) { 
                if (result === "1") {
                    callback();
                } else {
                    $('#eidvalidation').html('* The CV ID either does not exist or is not in the same language.');
                    $('#eidvalidation').show();
                    $('#modalContainer .modal-body .loaderContainer').remove();
                }
            },
            error : function (err) { 
                alert('[error]: ' + err.responseText); 
                $('#modalContainer .modal-body .loaderContainer').remove();
             }
        });
    }

    function profileCreation(initalCreate, callback) {
        var id = <%= iEid %>;
        var language = $('#languageToCreate').val();
        var existingId = $('#existingId').val();

        if (initalCreate === 0 && !(existingId.length > 0)) {
            $('#eidvalidation').html('* Please enter a valid expert id');
            $('#eidvalidation').show();
            return;
        }

        if (language === undefined || language === '') {
            language = $('#startingLanguage').val();
        }

        $('#modalContainer .modal-body').append('<div class="loaderContainer ml-25"><div style="text-align: center;" class="loader">loading...please wait...</div></div>')
        
        var url = database.includes("assortis") ? "expert_assortis_create_profile.asp" : "expert_create_profile.asp";
        var createProfileFunc = function () {
            $.ajax({
                url: url, 
                method: 'POST',
                data: { expertId: id, language: language, initalCreate: initalCreate, existingId: existingId !== '' ? existingId.split('-')[1] : null },
                success: callback, 
                error: function (err) { 
                    alert('[error]: ' + err.responseText); 
                    $('#modalContainer .modal-body .loaderContainer').remove();
                }
            })
        };

        if (initalCreate) {
            createProfileFunc();
            return;
        }
        
        var databaseId = existingId.split("-")[0];
        if (databaseId + "-" === "<%= sDatabaseId  %>") {
    
            checkIdAndLanguage(existingId.split("-")[1].trim(), language, database, createProfileFunc);

        }
        else {
            console.log("[Error]: different databases");
            alert('[error]: this CV is not in your database. Please contact the database owner.'); 
            $('#modalContainer .modal-body .loaderContainer').remove();
        }

        checkIdAndLanguage(existingId, language, database, createProfileFunc);
    }

    $('.btn').on('click', function () {
        $('.existingContainer').hide();
        var id = $(this).attr('data-id');

        $('.btn').removeClass('btn--selected');
        var item = $(this);
        if (!item.hasClass('btn--selected')) {
            item.addClass('btn--selected');
        }

        if (id === 'existing') {
            $('.existingContainer').show();
        }
        else {
            createNewProfile();
        }
    });
</script>
