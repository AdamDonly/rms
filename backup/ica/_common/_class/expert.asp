<%
Class CExpert
	Public ID
	Public Title
	Public FirstName
	Public MiddleName
	Public LastName

	Public BirthDate
	Public BirthPlace
	Public Gender
	Public MaritalStatus
	Public Comments

	Public Seniority
	Public Profession
	Public KeyExperience
	Public ProfessionalStatus
	
	Public Nationalities
	Public Educations
	Public Training
	Public Experiences
	Public Languages
	Public Addresses

	Public Default Property Get FullName
		FullName=Title & FirstName & MiddleName & LastName
	End Property

	Public Property Get TitleLastName
		FullName=Title & LastName
	End Property
	
	Public Function LoadDataFromRecordset(ARecordSet)
		If IsObject(ARecordSet) Then
			On Error Resume Next
			If Not ARecordSet.Eof Then
				ID=ARecordSet("id_Expert")

				FirstName=ARecordSet("psnFirstNameEng")
				MiddleName=ARecordSet("psnMiddleNameEng")
				LastName=ARecordSet("psnLastNameEng")

				BirthDate = objTempRs("psnBirthDate")
				BirthPlace = objTempRs("psnBirthPlaceEng")
				
				Gender=objTempRs("psnGender")
				MaritalStatus=objTempRs("id_MaritalStatus")

				Comments=objTempRs("expComments")
				
				Seniority=objTempRs("expProfYears")
				Profession=objTempRs("expProfessionEng")
				ProfessionalStatus=objTempRs("id_ProfessionalStatus")
				KeyQualification=objTempRs("expKeyQualificationsEng")
				
				CurentPosition=objTempRs("expCurrPositionEng")
				Memberships=objTempRs("expMemberProfEng")
				Publications=objTempRs("expPublicationsEng")
				References=objTempRs("expReferencesEng")
				
				Availability=objTempRs("expAvailabilityEng")
				Shortterm=objTempRs("expShortterm")
				Longterm=objTempRs("expLongterm")
				Preferences=objTempRs("expPreferences")
				
				Language=objTempRs("Lng")
				Email=objTempRs("Email")
				Phone=objTempRs("Phone")
				
			End If
			On Error GoTo 0
		End If
	End Function	
	
	Public Function LoadData(AProcedure, AParams)
	Dim objTempRs
		Set objTempRs=GetDataRecordsetSP(AProcedure, AParams)
		LoadDataFromRecordset objTempRs

		If objTempRs.State = adStateOpen Then objTempRs.Close
		Set objTempRs=Nothing
	End Function
	
End Class


%>
