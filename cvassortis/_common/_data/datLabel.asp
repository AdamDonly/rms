<%
Function GetLabel(ALanguage, ALabel)
	Dim sResult
	sResult = ""
	Dim sLanguage
	sLanguage = LCase(ALanguage)
	
	If sLanguage = "fr" _
	Or sLanguage = "fra" _
	Or sLanguage = "french" Then
		If dictLabel.Item("fr").Exists(ALabel) Then
			sResult = dictLabel.Item("fr").Item(ALabel)
		Else
			sResult = ALabel
		End If
	ElseIf sLanguage = "sp" _
	Or sLanguage = "es" _
	Or sLanguage = "spa" _
	Or sLanguage = "spanish" Then
		If dictLabel.Item("sp").Exists(ALabel) Then
			sResult = dictLabel.Item("sp").Item(ALabel)
		Else
			sResult = ALabel
		End If
	Else
		If dictLabel.Item("en").Exists(ALabel) Then
			sResult = dictLabel.Item("en").Item(ALabel)
		Else
			sResult = ALabel
		End If
	End If

GetLabel = sResult
End Function


Dim dictLabel
Set dictLabel = CreateObject("Scripting.Dictionary")

Dim dictLabelEnglish
Set dictLabelEnglish = CreateObject("Scripting.Dictionary")
Dim dictLabelFrench
Set dictLabelFrench = CreateObject("Scripting.Dictionary")
Dim dictLabelSpanish
Set dictLabelSpanish = CreateObject("Scripting.Dictionary")

dictLabel.Add "en", dictLabelEnglish
dictLabel.Add "fr", dictLabelFrench
dictLabel.Add "sp", dictLabelSpanish


' -----------------------------------------------------------------------------
' Login form

dictLabelFrench.Add "Login name", "Identifiant"
dictLabelSpanish.Add "Login name", "Nombre de Usuario"

dictLabelFrench.Add "Password", "Mot&nbsp;de&nbsp;passe"
dictLabelSpanish.Add "Password", "Clave"

dictLabelFrench.Add "Access is denied!", "Acc�dez est ni�!"
dictLabelSpanish.Add "Access is denied!", "Acceso denegado!"

dictLabelFrench.Add "Please fill in your login name", "Veuillez compl�ter votre nom d'ouverture"
dictLabelSpanish.Add "Please fill in your login name", "Por favor introduzca su nombre de usuario"

dictLabelFrench.Add "Please fill in your password", "Veuillez compl�ter votre mot de passe"
dictLabelSpanish.Add "Please fill in your password", "Por favor introduzca su clave"

dictLabelEnglish.Add "The login or password you supplied are not correct", "The login or password you supplied are not correct. They are case sensitive: check the CAPS LOCK key. Please verify the punctuation and spaces as well."
dictLabelFrench.Add "The login or password you supplied are not correct", "L'ouverture ou le mot de passe que vous avez fourni n'�tes pas correct. Ils distinguent les majuscules et minuscules : v�rifiez la clef de FONCTION MAJUSCULE. Veuillez v�rifier la ponctuation et les espaces aussi bien."
dictLabelSpanish.Add "The login or password you supplied are not correct", "El nombre de Usuario o clave que introdujo no es correcto. Verifique las mayusculas, puntuacion y espacios"

dictLabelFrench.Add "Please enter your login name and password", "Veuillez entrer votre identifiant et mot de passe"
dictLabelSpanish.Add "Please enter your login name and password", "Por favor introduzca su nombre de usuario y clave"

dictLabelFrench.Add "Forgot your login or password?", "Avez-vous oubli� votre Identifiant ou mot de passe?"
dictLabelSpanish.Add "Forgot your login or password?", "Olvido su nombre de usuario o clave?"

' -----------------------------------------------------------------------------
' Forgot password or login form

dictLabelFrench.Add "FORGOTTEN PASSWORD OR LOGIN", "MOT DE PASSE OU OUVERTURE OUBLI�"
dictLabelSpanish.Add "FORGOTTEN PASSWORD OR LOGIN", "NOMBRE DE USUARIO O CLAVE OLVIDADA"

dictLabelEnglish.Add "Your login and password has now been sent", "Your login and password has now been sent to your email. Please check your email and login."
dictLabelFrench.Add "Your login and password has now been sent", "Votre ouverture et mot de passe a �t� maintenant envoy�e � votre email. Veuillez v�rifier votre email et login."
dictLabelSpanish.Add "Your login and password has now been sent", "Su nombre de usuario y clave han sido enviadas a su e-mail. Por favor verifique su e-mail y nombre de usuario."

dictLabelEnglish.Add "The email address is not available", "The email address you supplied is not available in our database."
dictLabelFrench.Add "The email address is not available", "L'adresse email que vous avez fourni n'est pas disponible dans notre base de donn�es."
dictLabelSpanish.Add "The email address is not available", "La direccion de e-mail que nos facilito no esta disponible en nuestra base de datos."


' -----------------------------------------------------------------------------
' Homepage
dictLabelFrench.Add "Homepage", "Page d'accueil"
dictLabelSpanish.Add "Homepage", "Incio"

dictLabelFrench.Add "Search for experts", "Recherche des experts"
dictLabelSpanish.Add "Search for experts", "Buscar expertos"

dictLabelFrench.Add "Register new expert", "Encodage d'un nouvel expert"
dictLabelSpanish.Add "Register new expert", "Registrar nuevo experto"

dictLabelFrench.Add "Complete CV registration", "Enregistrer un nouveau CV"
dictLabelSpanish.Add "Complete CV registration", "Completar el registro de su CV"

dictLabelFrench.Add "Manage the database", "G�rez la base de donn�es"
dictLabelSpanish.Add "Manage the database", "Gestion de la base de datos"

dictLabelFrench.Add "List of all experts visible in the database", "Liste de tous les experts"
dictLabelSpanish.Add "List of all experts visible in the database", "Lista de todos los expertos en la base de datos"

dictLabelFrench.Add "List of experts registered this week", "Experts enregistr�s cette semaine"
dictLabelSpanish.Add "List of experts registered this week", "Expertos registrados esta semana" 

dictLabelFrench.Add "List of experts registered this month", "Experts enregistr�s ce mois"
dictLabelSpanish.Add "List of experts registered this month", "Expertos registrados este mes"

dictLabelFrench.Add "List of experts with CVs not updated for the past 12 months", "Experts n�ayant pas mis � jour leur CV pendant les 12 derniers mois"
dictLabelSpanish.Add "List of experts with CVs not updated for the past 12 months", "Lista de expertos cuyo cv no ha sido modificado en los ultimos 12 meses"

dictLabelFrench.Add "List of deleted experts", "Liste d'experts supprim�s"
dictLabelSpanish.Add "List of deleted experts", "Lista de expertos borrados"

dictLabelFrench.Add "Register new project", "Enregistrer un nouveau projet"
dictLabelSpanish.Add "Register new project", "Registrar nuevo proyecto"

dictLabelFrench.Add "New project", "Nouveau projet"
dictLabelSpanish.Add "New project", "Nuevo proyecto"

dictLabelFrench.Add "Projects. Tendering", "Projets soumissionn�s"
dictLabelSpanish.Add "Projects. Tendering", "Proyectos. Licitando"

dictLabelFrench.Add "Projects. Running", "Projets en cours"
dictLabelSpanish.Add "Projects. Running", "Proyectos. En curso"

dictLabelFrench.Add "Projects. Closed", "Projets achev�s"
dictLabelSpanish.Add "Projects. Closed", "Proyectos. Cerrado"

dictLabelFrench.Add "Projects. Inactive", "Projets inactifs"
dictLabelSpanish.Add "Projects. Inactive", "Proyectos. Inactivos"


' -----------------------------------------------------------------------------
' Project registration

dictLabelFrench.Add "PROJECT REGISTRATION", "ENREGISTREMENT DE PROJET"
dictLabelSpanish.Add "PROJECT REGISTRATION", "REGISTRAR PROYECTO"

dictLabelFrench.Add "Project status", "Statut de projet"
dictLabelSpanish.Add "Project status", "Estatus del proyecto"

dictLabelFrench.Add "Country / Region", "Pays / R�gion"
dictLabelSpanish.Add "Country / Region", "Pais / Region"

dictLabelFrench.Add "Description", "Description"
dictLabelSpanish.Add "Description", "Descripcion"

dictLabelFrench.Add "Deadline", "Date-limite"
dictLabelSpanish.Add "Deadline", "Fecha Limite"

' -----------------------------------------------------------------------------
' CV registration step 1 (register.asp)

dictLabelFrench.Add "CV registration",	"Enregistrement de CV"
dictLabelSpanish.Add "CV registration",	"Registrar CV"

dictLabelFrench.Add "Personal information",	"Informations personnelles"
dictLabelSpanish.Add "Personal information", "Informacion personal"

dictLabelEnglish.Add "If you have already registered your profile", "If you have already registered your profile, please <a href=""login.asp" & AddUrlParams(sParams, "url=" + sScriptFullName) & """>log in to update your details</a>."
dictLabelFrench.Add "If you have already registered your profile", "Si vous avez d�j� enregistr� votre profil, veuillez svp cliquer sur ce <a href=""login.asp" & AddUrlParams(sParams, "url=" + sScriptFullName) & """>lien pour mettre � jour votre CV</a>."
dictLabelSpanish.Add "If you have already registered your profile", "Si ya ha registrado su perfil, por favor <a href=""login.asp" & AddUrlParams(sParams, "url=" + sScriptFullName) & """> acceda para actualizar su CV</a>."

dictLabelEnglish.Add "Please fill in all the relevant information", "Please fill in all the relevant information and as many details on your experience as possible."
dictLabelFrench.Add "Please fill in all the relevant information", "Veuillez compl�ter toutes les informations importantes et autant de d�tails sur votre exp�rience que possible."
dictLabelSpanish.Add "Please fill in all the relevant information", "Por favor complete toda la informacion relevante y tantos detalles de su experiencia como sea posible."

dictLabelEnglish.Add "Fields marked with *", "Fields marked with <span class=""fcmp"">*</span> are required information."
dictLabelFrench.Add "Fields marked with *", "Les Champs identifi�s par <span class=""fcmp"">*</span> sont obligatoires."
dictLabelSpanish.Add "Fields marked with *", "Los campos identicifacos con <span class=""fcmp"">*</span> requiren ser completados."

dictLabelEnglish.Add "You can always go back", "You can always go back and edit each section by clicking on the menu at the top."
dictLabelFrench.Add "You can always go back", "Vous pouvez toujours retourner en arri�re et �diter chaque section en cliquant sur le menu au dessus."
dictLabelSpanish.Add "You can always go back", "Simepre puede retroceder y editar cada seccion pinchando en el menu situado arriba."

dictLabelFrench.Add "PERSONAL INFORMATION", "INFORMATIONS PERSONNELLES"
dictLabelSpanish.Add "PERSONAL INFORMATION", "INFORMACION PERSONAL"

dictLabelFrench.Add "CV language", "Langue du CV"
dictLabelSpanish.Add "CV language", "Lenguaje de CV"

dictLabelEnglish.Add "Personal title", "Title"
dictLabelFrench.Add "Personal title", "Civilit�"
dictLabelSpanish.Add "Personal title", "Titulo"

dictLabelFrench.Add "Please select", "Choisissez svp"
dictLabelSpanish.Add "Please select", "Pofavor eliga"

dictLabelFrench.Add "First name", "Pr�nom(s)"
dictLabelSpanish.Add "First name", "Nombre"

dictLabelFrench.Add "Middle name", "Deuxi�me pr�nom"
dictLabelSpanish.Add "Middle name", "Segundo nombre"

dictLabelFrench.Add "Family name", "Nom"
dictLabelSpanish.Add "Family name", "Apellido"

dictLabelFrench.Add "Last name", "Nom"
dictLabelSpanish.Add "Last name", "Apellido"

dictLabelFrench.Add "Surname", "Nom"
dictLabelSpanish.Add "Surname", "Apellido"

dictLabelFrench.Add "Name", "Nom"
dictLabelSpanish.Add "Name", "Nombre"

dictLabelFrench.Add "Surname(s) / First name(s)", "Nom(s) / Pr�nom(s)"
dictLabelSpanish.Add "Surname(s) / First name(s)", "Apellido(s) / Nombre(s)"

dictLabelEnglish.Add "Date of birth", "Date of birth"
dictLabelFrench.Add "Date of birth", "Date de naissance"
dictLabelSpanish.Add "Date of birth", "Fecha de nacimiento"

dictLabelFrench.Add "Day", "jour"
dictLabelSpanish.Add "Day", "dia"

dictLabelFrench.Add "Month", "mois"
dictLabelSpanish.Add "Month", "mes"

dictLabelFrench.Add "Year", "ann�e"
dictLabelSpanish.Add "Year", "a�o"

dictLabelEnglish.Add "Place of birth", "Place&nbsp;of&nbsp;birth"
dictLabelFrench.Add "Place of birth", "Lieu de naissance"
dictLabelSpanish.Add "Place of birth", "Lugar de naciemiento"

dictLabelFrench.Add "Civil status", "�tat civil"
dictLabelSpanish.Add "Civil status", "Estado civil"

dictLabelFrench.Add "Nationality", "Nationalit�"
dictLabelSpanish.Add "Nationality", "Nacionalidad"

dictLabelEnglish.Add "Add nationality", "Click on <b>Add</b> to add a selected nationality to your list."
dictLabelFrench.Add "Add nationality", "Cliquez sur <b>ajouter</b> pour ajouter une nationalit� choisie � votre liste."
dictLabelSpanish.Add "Add nationality", "Pinche en <b>ajouter</b> para incluir la nacionalidad elegido a su lista."

dictLabelEnglish.Add "Remove nationality", "If you want to remove a nationality, highlight it and click on <b>Remove</b>"
dictLabelFrench.Add "Remove nationality", "Si vous voulez supprimer une nationalit�, s�lectionnez-la et<br/>&nbsp; cliquez sur <b>supprimer</b>"
dictLabelSpanish.Add "Remove nationality", ""

dictLabelFrench.Add "Add", "Ajoutez"
dictLabelSpanish.Add "Add", "A�adir"

dictLabelFrench.Add "Remove", "Supprimer"
dictLabelSpanish.Add "Remove", "Suprimir"

dictLabelFrench.Add "Gender", "Sexe"
dictLabelSpanish.Add "Gender", "Sexo"

dictLabelFrench.Add "male", "masculin"
dictLabelSpanish.Add "male", "masculino"

dictLabelFrench.Add "female", "f�minin"
dictLabelSpanish.Add "female", "femenino"

dictLabelFrench.Add "Marital status", "�tat civil"
dictLabelSpanish.Add "Marital status", "Estado civil"

dictLabelFrench.Add "Primary phone", "T�l�phone"
dictLabelSpanish.Add "Primary phone", "Telefono principal"

dictLabelFrench.Add "Phone", "T�l�phone"
dictLabelSpanish.Add "Phone", "Telefono"

dictLabelFrench.Add "Primary email", "Email"
dictLabelSpanish.Add "Primary email", "Email principal"

dictLabelFrench.Add "Status", "Statut"
dictLabelSpanish.Add "Status", "Estado"

dictLabelFrench.Add "Current position", "Poste actuel"
dictLabelSpanish.Add "Current position", "Posicion actual"

dictLabelFrench.Add "Present position", "Situation pr�sente"
dictLabelSpanish.Add "Present position", "Situacion actual"

dictLabelFrench.Add "Desired employment / Occupational field", "Poste vis�"
dictLabelSpanish.Add "Desired employment / Occupational field", "Empleo deseado / campo de ocupacion"

dictLabelFrench.Add "Key qualifications", "Principales comp�tences"
dictLabelSpanish.Add "Key qualifications", "Competencias principales"

dictLabelFrench.Add "Years of professional experience", "Nombre d'ann�es d�exp�rience"
dictLabelSpanish.Add "Years of professional experience", "Numero de a�os de experiencia"

dictLabelFrench.Add "use only numbers", "utilisez uniquement des chiffres"
dictLabelSpanish.Add "use only numbers", "utilize solo numeros"

dictLabelFrench.Add "Specific experience in the region", "Exp�rience sp�cifique dans la r�gion"
dictLabelSpanish.Add "Specific experience in the region", "Experiencia especifica en la region"

dictLabelFrench.Add "Specific experience in non industrialised countries", "Exp�rience sp�cifique dans la r�gion"
dictLabelSpanish.Add "Specific experience in non industrialised countries", "Experiencia especifica en paises no industralizados"

dictLabelFrench.Add "SIRET number", "Num�ro SIRET"
dictLabelSpanish.Add "SIRET number", "Numero SIRET"

' -----------------------------------------------------------------------------
' CV registration step 2 (register2.asp)

dictLabelFrench.Add "Education", "�ducation"
dictLabelSpanish.Add "Education", "Educacion"

dictLabelFrench.Add "EDUCATION", "�DUCATION"
dictLabelSpanish.Add "EDUCATION", "EDUCACION"

dictLabelFrench.Add "EDUCATION AND TRAINING", "�DUCATION ET FORMATION"
dictLabelSpanish.Add "EDUCATION AND TRAINING", "EDUCACION Y FORMACION"

dictLabelFrench.Add "Education and training", "�ducation et formation"
dictLabelSpanish.Add "Education and training", "Educacion y formacion"

dictLabelFrench.Add "Please fill in the institution name", "Veuillez compl�ter le institution"
dictLabelSpanish.Add "Please fill in the institution name", "Por favor rellene la institucion"

dictLabelFrench.Add "Please fill in the education end date", "Veuillez compl�ter la date de fin de la formation"
dictLabelSpanish.Add "Please fill in the education end date", "Por favor rellene las fechas del fin de la formacion"

dictLabelFrench.Add "Please fill in the education dates properly", "Veuillez ins�rer les dates de formation correctement"
dictLabelSpanish.Add "Please fill in the education dates properly", "Por favor incluya los datos de formacion correctas"

dictLabelFrench.Add "Please specify a type of diploma or degree obtained", "Veuillez sp�cifier un type de dipl�me ou dipl�me universitire obtenu"
dictLabelSpanish.Add "Please specify a type of diploma or degree obtained", "Por favor especifique el tipo de diploma o carrera universitaria obtenido"

dictLabelFrench.Add "Please specify the education subject", "Veuillez sp�cifier les matieres d'enseignement"
dictLabelSpanish.Add "Please specify the education subject", "Por favor rellene el campo de educacion"

dictLabelFrench.Add "No.", "Num�ro"
dictLabelSpanish.Add "No.", "Numero"

dictLabelEnglish.Add "Institution name", "Institution&nbsp;name"
dictLabelFrench.Add "Institution name", "Institution&nbsp;o�&nbsp;le<br />dipl�me&nbsp;a&nbsp;�t�&nbsp;obtenu"
dictLabelSpanish.Add "Institution name", "Nombre&nbsp;de&nbsp;institucion"

dictLabelFrench.Add "Institution", "Institution"
dictLabelSpanish.Add "Institution", "Institucion"

dictLabelFrench.Add "Start date", "Date&nbsp;de&nbsp;d�but"
dictLabelSpanish.Add "Start date", "Fecha de comienzo"

dictLabelFrench.Add "Date from", "Date d�but"
dictLabelSpanish.Add "Date from", "Fecha de"

dictLabelFrench.Add "End date", "Date&nbsp;de&nbsp;fin"
dictLabelSpanish.Add "End date", "Fecha de fin"

dictLabelFrench.Add "Date to", "Date fin"
dictLabelSpanish.Add "Date to", "Fecha a"

dictLabelFrench.Add "Dates (from � to)", "Date (d�but - fin)"
dictLabelSpanish.Add "Dates (from � to)", "Fechas (comienzo - fin)"

dictLabelFrench.Add "Subject", "Sujet"
dictLabelSpanish.Add "Subject", "Campo"

dictLabelFrench.Add "Modify", "Modifiez"
dictLabelSpanish.Add "Modify", "Modificar"

dictLabelFrench.Add "Delete", "Supprimez"
dictLabelSpanish.Add "Delete", "Borrar"

dictLabelEnglish.Add "Type of diploma", "Type of Diploma /<br />Degree obtained"
dictLabelFrench.Add "Type of diploma", "Type de dipl�me obtenu"
dictLabelSpanish.Add "Type of diploma", "Tipo de diploma obtenido"

dictLabelFrench.Add "Degree(s) or Diploma(s) obtained", "Dipl�me(s) obtenu(s)"
dictLabelSpanish.Add "Degree(s) or Diploma(s) obtained", "Diplomas obtenidos"

dictLabelFrench.Add "If other please specify", "Si autre veuillez sp�cifier"
dictLabelSpanish.Add "If other please specify", "Si otro, por favor especifique"

dictLabelFrench.Add "If needed, please specify the exact title of your diploma", "Si n�cessaire sp�cifiez le titre exact de votre dipl�me"
dictLabelSpanish.Add "If needed, please specify the exact title of your diploma", "Por favor especifique el titulo exacto de su diploma"

dictLabelFrench.Add "If needed, please specify the exact title of your degree", "Si n�cessaire sp�cifiez le titre exact de votre dipl�me"
dictLabelSpanish.Add "If needed, please specify the exact title of your degree", "Por favor especifique el titulo exacto de su diploma"

dictLabelFrench.Add "Exact title of your degree", "Le titre exact de votre dipl�me"
dictLabelSpanish.Add "Exact title of your degree", "El titulo exacto de su diploma"

dictLabelFrench.Add "Name and type of organisation providing education and training", "Nom et type de l'�tablissement d'enseignement ou de formation"
dictLabelSpanish.Add "Name and type of organisation providing education and training", "Nombre y tipo de establecimiento de formacion"

dictLabelFrench.Add "Principal subjects/occupational skills covered", "Principales mati�res/comp�tences professionnelles couvertes"
dictLabelSpanish.Add "Principal subjects/occupational skills covered", "Principales temas cubiertos"

dictLabelFrench.Add "Title of qualification awarded", "Intitul� du certificat ou dipl�me d�livr�"
dictLabelSpanish.Add "Title of qualification awarded", "Titulo de diploma obtenido"

dictLabelFrench.Add "Level in national or international classification", "Niveau dans la classification nationale ou internationale"
dictLabelSpanish.Add "Level in national or international classification", "Nivel en la clasificacion nacional o internacional"


' -----------------------------------------------------------------------------
' CV registration step 2 (register21.asp)

dictLabelFrench.Add "Training", "Autre formation"
dictLabelSpanish.Add "Training", "Otra formacion"

dictLabelFrench.Add "TRAINING", "AUTRE FORMATION"
dictLabelSpanish.Add "TRAINING", "OTRA FORMACION"

dictLabelFrench.Add "Please fill in the training title", "Veuillez compl�ter le titre de formation"
dictLabelSpanish.Add "Please fill in the training title", "Por favor rellene el titulo de la formacion"

dictLabelFrench.Add "Please fill in the training end date", "Veuillez ins�rer la date de fin de formation"
dictLabelSpanish.Add "Please fill in the training end date", "Por favor rellene la fecha de fin de la formacion"

dictLabelFrench.Add "Please fill in the training dates properly", "Veuillez ins�rer les dates de formation correctement"
dictLabelSpanish.Add "Please fill in the training dates properly", "Por favor, agregue la informacion correcta"

dictLabelFrench.Add "Not specified", "Non sp�cifi�"
dictLabelSpanish.Add "Not specified", "Sin especificar"

dictLabelFrench.Add "Title", "Titre"
dictLabelSpanish.Add "Title", "Titulo"

dictLabelFrench.Add "Skills / Qualifications", "Qualifications"
dictLabelSpanish.Add "Skills / Qualifications", "Cualificaciones"

dictLabelFrench.Add "Achievements", "Accomplissements"
dictLabelSpanish.Add "Achievements", "Logros"


' -----------------------------------------------------------------------------
' CV registration step 3 (register3.asp)

dictLabelFrench.Add "Professional experience", "Exp�riences professionnelles"

dictLabelFrench.Add "Please specify the project title or the name of the company or organisation", "Veuillez sp�cifier le titre du projet ou le nom de la compagnie ou de l'organisation"

dictLabelFrench.Add "Please fill in the experience start date", "Veuillez ins�rer la date de d�but d'exp�rience"

dictLabelFrench.Add "Please fill in the experience end date", "Veuillez ins�rer la date de fin d'exp�rience"

dictLabelFrench.Add "Please fill in the experience dates properly", "Veuillez ins�rer les dates d'exp�rience correctement"

dictLabelFrench.Add "Please fill in your position", "Veuillez compl�ter votre position"

dictLabelFrench.Add "Please make the description of the project shorter", "Veuillez rendre la description du projet plus courte"

dictLabelFrench.Add "Please select at least one country", "Veuillez choisir au moins un pays"

dictLabelFrench.Add "You cannot select more than 30 countries for one project", "Vous ne pouvez pas choisir plus de 30 pays pour un projet"

dictLabelFrench.Add "Please select at least one sub-sector of expertise", "Veuillez choisir au moins un sous-secteur d'expertise"

dictLabelFrench.Add "You cannot select more than 50 sectors for one project", "Vous ne pouvez pas choisir plus de 50 secteurs pour un projet"

dictLabelFrench.Add "Project title", "Titre du projet"

dictLabelEnglish.Add "Type of experience (Reg)", "Type of experience<br /><small>(if relevant)</small>"
dictLabelFrench.Add "Type of experience (Reg)", "Type d'exp�rience<br /><small>(le cas �ch�ant)</small>"

dictLabelEnglish.Add "Project title (Reg)", "Project title<br /><small>(if relevant)</small>"
dictLabelFrench.Add "Project title (Reg)", "Titre du projet<br /><small>(le cas �ch�ant)</small>"

dictLabelEnglish.Add "Main project features (Reg)", "Main project features<br /><small>(if relevant)</small>"
dictLabelFrench.Add "Main project features (Reg)", "Caract�ristiques principales du projet<br /><small>(le cas �ch�ant)</small>"

dictLabelFrench.Add "Position", "Position"

dictLabelFrench.Add "Project / Organisation", "Projet / organisation"

dictLabelFrench.Add "Company / Organisation", "Compagnie / organisation"

dictLabelFrench.Add "Position / Responsibility", "Position / responsabilit�"

dictLabelFrench.Add "Beneficiary", "B�n�ficiaire"

dictLabelFrench.Add "Location", "Lieu"

dictLabelFrench.Add "Countries", "Pays"

dictLabelFrench.Add "Sectors", "Secteurs"

dictLabelFrench.Add "Client references", "R�f�rences&nbsp;du&nbsp;client"

dictLabelFrench.Add "Company and reference person", "Soci�t� et personne de r�f�rence"

dictLabelEnglish.Add "Brief description of tasks", "Brief description of<br />the tasks assigned"
dictLabelFrench.Add "Brief description of tasks", "Courte description <br />des t�ches assign�es"

dictLabelEnglish.Add "Description of tasks", "Description of<br />the tasks assigned"
dictLabelFrench.Add "Description of tasks", "Description <br />des t�ches assign�es"

dictLabelFrench.Add "Funding agency", "Agence&nbsp;de&nbsp;placement"

dictLabelFrench.Add "Major funding agencies", "Principaux organismes de financement"

dictLabelFrench.Add "Other funding agencies", "Autres organismes de financement"

dictLabelEnglish.Add "Select funding agency from list", "Select funding agency from the list or specify in the field above if it is not in the list"
dictLabelFrench.Add "Select funding agency from list", "Choisissez l'agence de placement � partir de la liste ou sp�cifiez dans le domaine ci-dessus s'il n'est pas dans la liste"

dictLabelFrench.Add "SELECT PROJECT'S COUNTRIES", "CHOISISSEZ LES PAYS DU PROJET"

dictLabelFrench.Add "SELECT PROJECT'S SUB-SECTORS", "CHOISISSEZ LES SOUS-SECTEURS DU PROJET"

dictLabelEnglish.Add "KEY QUALIFICATION AND SPECIFIC EXPERIENCE", "KEY QUALIFICATION AND SPECIFIC EXPERIENCE (PROJECTS, ETC.)"
dictLabelFrench.Add "KEY QUALIFICATION AND SPECIFIC EXPERIENCE", "PRINCIPALES QUALIFICATIONS ET EXPERIENCES SP�CIFIQUE (PROJETS, ETC)"

dictLabelEnglish.Add "Key qualification and specific experience", "Key qualification and specific experience (projects, etc.)"
dictLabelFrench.Add "Key qualification and specific experience", "Principales qualifications et experiences sp�cifique (projets, etc)"

dictLabelFrench.Add "PROFESSIONAL EXPERIENCE", "EXP�RIENCES PROFESSIONNELLES"

dictLabelFrench.Add "EMPLOYMENT RECORD AND COMPLETED PROJECTS", "EMPLOIS PASSES ET PROJETS R�ALIS�S"
dictLabelFrench.Add "Employment record and completed projects", "Emplois passes et projets r�alis�s"

dictLabelFrench.Add "WORK EXPERIENCE", "EXP�RIENCE PROFESSIONNELLE"

dictLabelFrench.Add "Work experience", "Exp�rience professionnelle"

dictLabelFrench.Add "Ongoing", "En cours"

dictLabelFrench.Add "Reference", "R�f�rence"

dictLabelFrench.Add "contact person", "personne � contacter"

dictLabelFrench.Add "Occupation or position held", "Fonction ou poste occup�"

dictLabelFrench.Add "Main activities and responsibilities", "Principales activit�s et responsabilit�s"

dictLabelFrench.Add "Name and address of employer", "Nom et adresse de l'employeur"

dictLabelFrench.Add "Type of business or sector", "Type ou secteur d�activit�"

dictLabelEnglish.Add "Specify professional experiences (GIP)", "Specify your main professional experiences in France or abroad.<br />Fill in the complete form for each experience, starting with the most recent one. Once you have filled in the form for one experience, click on [ Add an Experience ] to add the information related to your previous experiences."
dictLabelFrench.Add "Specify professional experiences (GIP)", "Sp�cifiez vos principales exp�riences en France ou ailleurs.<br />Remplissez le formulaire complet pour chaque exp�rience, commen�ant par la plus r�cente. Une fois vouz auriez rempli le formulaire pour une exp�rience, clicquer sur [ Ajouter une exp�rience ] pour ajouter l'information concernant vos exp�riences ant�rieures."

dictLabelEnglish.Add "SELECT PROJECT'S COUNTRIES (GIP)", "SELECT COUNTRIES IN WHICH <br/>&nbsp; &nbsp; &nbsp; &nbsp; YOU HAVE WORKED DURING THIS EXPERIENCE"
dictLabelFrench.Add "SELECT PROJECT'S COUNTRIES (GIP)", "CHOISISSEZ LE OU LES PAYS OU VOUS AVEZ <br/>&nbsp; &nbsp; &nbsp; &nbsp; ACCOMPLI CETTE EXPERIENCE/MISSION"

dictLabelFrench.Add "only for international projects", "uniquement pour les projets internationaux"

dictLabelEnglish.Add "First click on the sector title (GIP)", "First click on the sector title in the left column and sub-sectors will appear in the right column.<br />Then select the sub-sectors that best match your experience."
dictLabelFrench.Add "First click on the sector title (GIP)", "Cliquez d�abord sur le secteur, dans la colonne de gauche et les sous-secteurs appara�tront dans la colonne de droite. Choisissez alors les sous-secteurs qui correspondent au mieux � l�exp�rience que vous d�crivez. "

' -----------------------------------------------------------------------------
' CV registration step 4 (register4.asp)

dictLabelFrench.Add "Languages", "Langues"

dictLabelFrench.Add "Please select your native language", "S'il vous pla�t s�lectionnez votre langue maternelle"

dictLabelFrench.Add "Please choose a language", "S'il vous pla�t choisissez une langue"

dictLabelFrench.Add "Please choose the levels of your knowledge", "S'il vous pla�t choisissez le niveau de vos connaissances"

dictLabelFrench.Add "You are only allowed 20 languages", "Vous �tes seulement autoris� 20 langues"

dictLabelEnglish.Add "Add selected language", "Click on [ Add ] button to add a selected language to your list.<br />If you want to remove a language, highlight it and click on [ Remove ] button."
dictLabelFrench.Add "Add selected language", "S�lectionnez votre ou vos langues maternelles et cliques sur [ Ajouter ] pour les ajouter � la liste. Si vous souhaitez supprimer une langue, s�lectionnez-la et cliquez sur [ Enlever ]."

dictLabelEnglish.Add "Choose a language and specify your level...", "Choose a language and specify your level (reading, speaking, writing). To add another languages click on [ Add language ] button and specify the levels for each of them."
dictLabelFrench.Add "Choose a language and specify your level...", "Choisissez une langue et sp�cifiez votre niveau (lu, parl�, �crit). Pour ajouter une langue cliquez sur [ Ajouter une langue ] et sp�cifiez � chaque fois les niveaux pour chacune d'entre elles."

dictLabelFrench.Add "Language", "Langue"

dictLabelFrench.Add "Native", "langue maternelle"

dictLabelFrench.Add "Mother tongue(s)", "Langue(s) maternelle(s)"

dictLabelFrench.Add "Other language(s)", "Autre(s) langue(s)"

dictLabelFrench.Add "Reading", "Lu"

dictLabelEnglish.Add "Reading(EP)", "Reading"
dictLabelFrench.Add "Reading(EP)", "Lire"

dictLabelFrench.Add "Understanding", "Comprendre"

dictLabelFrench.Add "Speaking", "Parl�"

dictLabelEnglish.Add "Speaking(EP)", "Speaking"
dictLabelFrench.Add "Speaking(EP)", "Parler"

dictLabelFrench.Add "Listening", "�couter"

dictLabelFrench.Add "Writing", "Ecrit"

dictLabelEnglish.Add "Writing(EP)", "Writing"
dictLabelFrench.Add "Writing(EP)", "�crire"

dictLabelFrench.Add "NATIVE LANGUAGES", "LANGUE MATERNELLES"
dictLabelFrench.Add "Native languages", "Langue maternelles"

dictLabelFrench.Add "FOREIGN LANGUAGES", "LANGUES �TRANG�RES"
dictLabelFrench.Add "Foreign languages", "Langues �trang�res"

dictLabelFrench.Add "LANGUAGES SKILLS", "CAPACIT�S LINGUISTIQUES"
dictLabelFrench.Add "Languages skills", "Capacit�s linguistiques"

dictLabelEnglish.Add "Languages skills EC", "Language skills: Indicate competence on a scale of 1 to 5 (1 - excellent; 5 - basic)"
dictLabelFrench.Add "Languages skills EC", "Connaissances linguistiques: Indiquer vos connaissances sur une �chelle de 1 � 5 (1 - niveau excellent; 5 - niveau rudimentaire)"

dictLabelFrench.Add "Common European Framework of Reference for Languages", "Cadre europ�en commun de r�f�rence pour les langues"

dictLabelFrench.Add "Spoken interaction", "Prendre part � une conversation"

dictLabelFrench.Add "Spoken production", "S�exprimer oralement en continu"

' -----------------------------------------------------------------------------
' CV registration step 5 (register5.asp)

dictLabelEnglish.Add "Contact details & availability", "Contact&nbsp;details<br />&amp;&nbsp;availability"
dictLabelFrench.Add "Contact details & availability", "Coordonn�es<br />&amp;&nbsp;disponibilit�"

dictLabelFrench.Add "Contact details", "Coordonn�es"

dictLabelFrench.Add "availability", "disponibilit�"

dictLabelFrench.Add "permanent address", "adresse permanente"

dictLabelFrench.Add "Please fill in a street of ", "S'il vous pla�t remplir la rue de "

dictLabelFrench.Add "Please fill in a city of ", "S'il vous pla�t remplir la ville de "

dictLabelFrench.Add "Please fill in a postcode of ", "S'il vous pla�t remplir le code postal de "

dictLabelFrench.Add "Please select a country of ", "S'il vous pla�t s�lectionner le pays de "

dictLabelFrench.Add "Please fill in ", "S'il vous pla�t remplir "

dictLabelFrench.Add " permanent phone number.", " num�ro de t�l�phone."

dictLabelFrench.Add "Please specify ", "S'il vous pla�t pr�cisez "

dictLabelFrench.Add " permanent email.", " adresse email."

dictLabelFrench.Add "Please retype ", "S'il vous pla�t retaper "

dictLabelFrench.Add "permanent email correctly", "adresse email correctement"

dictLabelFrench.Add "current email correctly", "actuel courriel correctement"

dictLabelFrench.Add "Please make text of your availibility shorter.", "S'il vous pla�t rendre le texte de votre disponibilit� plus court."

dictLabelFrench.Add "Street", "Rue"

dictLabelFrench.Add "City", "Ville"

dictLabelFrench.Add "Postcode", "Code postal"

dictLabelFrench.Add "Country", "Pays"

dictLabelFrench.Add "Mobile", "GSM"

dictLabelFrench.Add "Fax", "Fax"

dictLabelFrench.Add "Website", "Site Web"

dictLabelEnglish.Add "Please specify availability", "Please specify the periods in which you are available in the next two years.<br />To guarantee the best matches, please keep your availability information updated."
dictLabelFrench.Add "Please specify availability", "Veuillez sp�cifier les p�riodes durant lesquelles vous �tes disponible pendant les deux ann�es � venir. Merci de maintenir � jour les informations relatives � vos disponibilit�s."

dictLabelFrench.Add "Availability", "Disponibilit�"

dictLabelFrench.Add "Availability & preferences", "Disponibilit� & pr�f�rences"

dictLabelFrench.Add "location", "pays"

dictLabelEnglish.Add "Please state your preferences", "Please state your preferences for short-and/or long-term missions."
dictLabelFrench.Add "Please state your preferences", "Veuillez sp�cifier vos pr�f�rences pour des missions courtes et / ou de long terme."

dictLabelFrench.Add "Preferences", "Pr�f�rences"

dictLabelFrench.Add "Short-term missions", "Missions � court terme"

dictLabelFrench.Add "Long-term missions", "Missions � long terme"

dictLabelFrench.Add "Other skills", "Autres comp�tences"

dictLabelFrench.Add "Other preferences", "D'autres pr�f�rences"

dictLabelFrench.Add "Membership in professional bodies", "Adh�sion � des associations ou corps professionnels"

dictLabelFrench.Add "Membership of professional bodies", "Affiliation � une organisation professionnelle"

dictLabelFrench.Add "Publications", "Publications"

dictLabelFrench.Add "References", "R�f�rences"

dictLabelFrench.Add "PERMANENT ADDRESS", "ADRESSE PERMANENTE"
dictLabelFrench.Add "Permanent address", "Adresse permanente"

dictLabelFrench.Add "CURRENT ADDRESS", "ADRESSE ACTUELLE"
dictLabelFrench.Add "Current address", "Adresse actuelle"

dictLabelFrench.Add "CURRENT ADDRESS (if different)", "ADRESSE ACTUELLE (SI ELLE EST DIFF�RENTE)"
dictLabelFrench.Add "Current address (if different)", "Adresse actuelle (si elle est diff�rente)"

dictLabelFrench.Add "CURRENT AVAILABILITY", "DISPONIBILIT� ACTUELLE"
dictLabelFrench.Add "Current availability", "Disponibilit� actuelle"

dictLabelFrench.Add "MISCELLANEOUS", "DIVERS"
dictLabelFrench.Add "Miscellaneous", "Divers"

dictLabelFrench.Add "ASSIGNMENT PREFERENCES", "PR�F�RENCE DES MISSIONS"
dictLabelFrench.Add "Assignment preferences", "Pr�f�rences des missions"

dictLabelFrench.Add "OTHER", "AUTRE"
dictLabelFrench.Add "Other", "Autre"

dictLabelFrench.Add "Address", "Adresse"

dictLabelFrench.Add "Other relevant information (e.g., Publications)", "Autres informations pertinentes (p,ex., r�f�rences de publications)"

dictLabelFrench.Add "Personal skills and competences", "Aptitudes et comp�tences personnelles"

dictLabelFrench.Add "Social skills and competences", "Aptitudes et comp�tences sociales"

dictLabelFrench.Add "Organisational skills and competences", "Aptitudes et comp�tences organisationnelles"

dictLabelFrench.Add "Technical skills and competences", "Aptitudes et comp�tences techniques"

dictLabelFrench.Add "Computer skills and competences", "Aptitudes et comp�tences informatiques"

dictLabelFrench.Add "Artistic skills and competences", "Aptitudes et comp�tences artistiques"

dictLabelFrench.Add "Other skills and competences", "Autres aptitudes et comp�tences"

dictLabelFrench.Add "Driving licence(s)", "Permis de conduire"

dictLabelFrench.Add "Additional information", "Information compl�mentaire"

' -----------------------------------------------------------------------------
' CV WB format 

dictLabelFrench.Add "FORM TECH-6", "FORMULAIRE TECH-6"
dictLabelFrench.Add "CURRICULUM VITAE (CV) FOR PROPOSED PROFESSIONAL STAFF", "CURRICULUM VITAE (CV) DU PERSONNEL CLE PROPOSE"

dictLabelFrench.Add "Proposed Position", "Poste"

dictLabelEnglish.Add "only one candidate...", "only one candidate shall be nominated for each position"
dictLabelFrench.Add "only one candidate...", "un seul candidat par poste"

dictLabelFrench.Add "Name of Firm", "Nom du consultant"

dictLabelFrench.Add "insert name of firm proposing the staff", "indiquer le nom de la soci�t� proposant le personnel"

dictLabelFrench.Add "Name of Staff", "Nom de l�employ�"

dictLabelFrench.Add "Date of Birth", "Date de naissance"

dictLabelEnglish.Add "Education (WB)", "Education"
dictLabelFrench.Add "Education (WB)", "Formation"

dictLabelFrench.Add "Name of assignment or project", "Nom du projet ou de la mission"

dictLabelFrench.Add "Main project features", "Principales caract�ristiques du projet"

dictLabelEnglish.Add "Indicate college/university...", "Indicate college/university and other specialized education of staff member, giving names of institutions, degrees obtained, and dates of obtainment"
dictLabelFrench.Add "Indicate college/university...", "Indiquer les �tudes universitaires et autres �tudes sp�cialis�es de l�employ� ainsi que les noms des institutions fr�quent�es, les dipl�mes obtenus et les dates auxquelles ils l�ont �t�"

dictLabelFrench.Add "Membership of Professional Associations", "Affiliation � des associations/groupements professionnels"

dictLabelFrench.Add "Membership in Professional Societies", "Affiliation � des associations professionnelles"

dictLabelFrench.Add "Other Training", "Autres formations"

dictLabelEnglish.Add "Indicate significant training...", "Indicate significant training since degrees under 5 - Education were obtained"
dictLabelFrench.Add "Indicate significant training...", "Indiquer toute autre formation re�ue depuis 5 ci-dessus"

dictLabelFrench.Add "Countries of Work Experience", "Pays o� l�employ� a travaill�"

dictLabelFrench.Add "List countries where staff has worked in the last ten years", "Donner la liste des pays ou l�employ� a travaill� au cours des 10 derni�res ann�es"

dictLabelEnglish.Add "For each language indicate proficiency...", "For each language indicate proficiency: good, fair, or poor in speaking, reading, and writing"
dictLabelFrench.Add "For each language indicate proficiency...", "Indiquer pour chacune le degr� de connaissance : bon, moyen, m�diocre pour ce qui est de la langue parl�e, lue et �crite"

dictLabelFrench.Add "Employment Record", "Exp�rience professionnelle"

dictLabelEnglish.Add "Starting with present position...", "Starting with present position, list in reverse order every employment held by staff member since graduation, giving for each employment (see format here below): dates of employment, name of employing organization, positions held"
dictLabelFrench.Add "Starting with present position...", "En commen�ant par son poste actuel, donner la liste par ordre chronologique inverse de tous les emplois exerc�s par l�employ� depuis la fin de ses �tudes. Pour chaque emploi (voir le formulaire ci-dessous), donner les dates, le nom de l�employeur et le poste occup�."

'dictLabelFrench.Add "", "Depuis [ann�e]"

'dictLabelFrench.Add "", "jusqu�� [ann�e]"

dictLabelFrench.Add "Employer", "Employeur"

dictLabelFrench.Add "Position held", "Poste"

dictLabelFrench.Add "Activities performed", "Activit�s r�alis�es"

dictLabelFrench.Add "Description of duties", "Description des t�ches assign�es"

dictLabelFrench.Add "Detailed Tasks Assigned", "D�tail des t�ches ex�cut�es"

dictLabelFrench.Add "Detailed Tasks Assigned (AFB)", "Attributions sp�cifiques"

dictLabelFrench.Add "List all tasks to be performed under this assignment", "Indiquer toutes les t�ches � ex�cuter dans le cadre de cette proposition"

dictLabelFrench.Add "Work Undertaken that Best Illustrates Capability to Handle the Tasks Assigned", "Exp�rience de l�employ� qui illustre le mieux sa comp�tence"

dictLabelEnglish.Add "Among the assignments in which the staffs have been involved...", "Among the assignments in which the staffs have been involved, indicate the following information for those assignments that best illustrate staff capability to handle the tasks listed under point 11."
dictLabelFrench.Add "Among the assignments in which the staffs have been involved...", "Donner notamment les informations suivantes qui illustrent au mieux la comp�tence professionnelle de l�employ� pour les t�ches mentionn�es au point 11."

dictLabelEnglish.Add "Among the assignments in which the staffs have been involved (ADB)...", "Among the assignments in which the staffs have been involved, indicate the following information for those assignments that best illustrate staff capability to handle the tasks listed under point 12."
dictLabelFrench.Add "Among the assignments in which the staffs have been involved (ADB)...", "Donner notamment les informations suivantes qui illustrent au mieux la comp�tence professionnelle de l�employ� pour les t�ches mentionn�es au point 12."

dictLabelFrench.Add "Certification", "Attestation"

dictLabelEnglish.Add "I, the undersigned, certify...", "I, the undersigned, certify that to the best of my knowledge and belief, this CV correctly describes me, my qualifications, and my experience.  I understand that any wilful misstatement described herein may lead to my disqualification or dismissal, if engaged."
dictLabelFrench.Add "I, the undersigned, certify...", "Je, soussign�, certifie, en toute conscience, que les renseignements ci-dessus rendent fid�lement compte de ma situation, de mes qualifications et de mon exp�rience. J�accepte que toute d�claration volontairement erron�e peut entra�ner mon exclusion, ou mon renvoi si j�ai �t� engag�."

dictLabelEnglish.Add "I, the undersigned, certify (AFB)...", "I, the undersigned, certify that to the best of my knowledge and belief, these biodata correctly describe myself, my qualifications and my experience."
dictLabelFrench.Add "I, the undersigned, certify (AFB)...", "Je, soussign�, certifie, sur la base des donn�es � ma disposition, que les renseignements ci-dessus rendent fid�lement compte de ma situation, de mes qualifications et de mon exp�rience."

dictLabelFrench.Add "Signature of staff member or authorized representative of the staff", "Signature de l�employ� et du repr�sentant habilit� du consultant"

dictLabelFrench.Add "Day/Month/Year", "Jour/mois/ann�e"

dictLabelFrench.Add "Full name of authorized representative", "Nom du repr�sentant habilit�"

dictLabelFrench.Add "Years with Firm", "Ann�es d�emploi au sein de la firme"

dictLabelEnglish.Add "Key Qualifications (AFD)", "Key Qualifications"
dictLabelFrench.Add "Key Qualifications (AFD)", "Principales qualifications"
%>