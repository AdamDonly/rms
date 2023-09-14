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

dictLabelFrench.Add "Access is denied!", "Accédez est nié!"
dictLabelSpanish.Add "Access is denied!", "Acceso denegado!"

dictLabelFrench.Add "Please fill in your login name", "Veuillez compléter votre nom d'ouverture"
dictLabelSpanish.Add "Please fill in your login name", "Por favor introduzca su nombre de usuario"

dictLabelFrench.Add "Please fill in your password", "Veuillez compléter votre mot de passe"
dictLabelSpanish.Add "Please fill in your password", "Por favor introduzca su clave"

dictLabelEnglish.Add "The login or password you supplied are not correct", "The login or password you supplied are not correct. They are case sensitive: check the CAPS LOCK key. Please verify the punctuation and spaces as well."
dictLabelFrench.Add "The login or password you supplied are not correct", "L'ouverture ou le mot de passe que vous avez fourni n'êtes pas correct. Ils distinguent les majuscules et minuscules : vérifiez la clef de FONCTION MAJUSCULE. Veuillez vérifier la ponctuation et les espaces aussi bien."
dictLabelSpanish.Add "The login or password you supplied are not correct", "El nombre de Usuario o clave que introdujo no es correcto. Verifique las mayusculas, puntuacion y espacios"

dictLabelFrench.Add "Please enter your login name and password", "Veuillez entrer votre identifiant et mot de passe"
dictLabelSpanish.Add "Please enter your login name and password", "Por favor introduzca su nombre de usuario y clave"

dictLabelFrench.Add "Forgot your login or password?", "Avez-vous oublié votre Identifiant ou mot de passe?"
dictLabelSpanish.Add "Forgot your login or password?", "Olvido su nombre de usuario o clave?"

' -----------------------------------------------------------------------------
' Forgot password or login form

dictLabelFrench.Add "FORGOTTEN PASSWORD OR LOGIN", "MOT DE PASSE OU OUVERTURE OUBLIÉ"
dictLabelSpanish.Add "FORGOTTEN PASSWORD OR LOGIN", "NOMBRE DE USUARIO O CLAVE OLVIDADA"

dictLabelEnglish.Add "Your login and password has now been sent", "Your login and password has now been sent to your email. Please check your email and login."
dictLabelFrench.Add "Your login and password has now been sent", "Votre ouverture et mot de passe a été maintenant envoyée à votre email. Veuillez vérifier votre email et login."
dictLabelSpanish.Add "Your login and password has now been sent", "Su nombre de usuario y clave han sido enviadas a su e-mail. Por favor verifique su e-mail y nombre de usuario."

dictLabelEnglish.Add "The email address is not available", "The email address you supplied is not available in our database."
dictLabelFrench.Add "The email address is not available", "L'adresse email que vous avez fourni n'est pas disponible dans notre base de données."
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

dictLabelFrench.Add "Manage the database", "Gérez la base de données"
dictLabelSpanish.Add "Manage the database", "Gestion de la base de datos"

dictLabelFrench.Add "List of all experts visible in the database", "Liste de tous les experts"
dictLabelSpanish.Add "List of all experts visible in the database", "Lista de todos los expertos en la base de datos"

dictLabelFrench.Add "List of experts registered this week", "Experts enregistrés cette semaine"
dictLabelSpanish.Add "List of experts registered this week", "Expertos registrados esta semana" 

dictLabelFrench.Add "List of experts registered this month", "Experts enregistrés ce mois"
dictLabelSpanish.Add "List of experts registered this month", "Expertos registrados este mes"

dictLabelFrench.Add "List of experts with CVs not updated for the past 12 months", "Experts n’ayant pas mis à jour leur CV pendant les 12 derniers mois"
dictLabelSpanish.Add "List of experts with CVs not updated for the past 12 months", "Lista de expertos cuyo cv no ha sido modificado en los ultimos 12 meses"

dictLabelFrench.Add "List of deleted experts", "Liste d'experts supprimés"
dictLabelSpanish.Add "List of deleted experts", "Lista de expertos borrados"

dictLabelFrench.Add "Register new project", "Enregistrer un nouveau projet"
dictLabelSpanish.Add "Register new project", "Registrar nuevo proyecto"

dictLabelFrench.Add "New project", "Nouveau projet"
dictLabelSpanish.Add "New project", "Nuevo proyecto"

dictLabelFrench.Add "Projects. Tendering", "Projets soumissionnés"
dictLabelSpanish.Add "Projects. Tendering", "Proyectos. Licitando"

dictLabelFrench.Add "Projects. Running", "Projets en cours"
dictLabelSpanish.Add "Projects. Running", "Proyectos. En curso"

dictLabelFrench.Add "Projects. Closed", "Projets achevés"
dictLabelSpanish.Add "Projects. Closed", "Proyectos. Cerrado"

dictLabelFrench.Add "Projects. Inactive", "Projets inactifs"
dictLabelSpanish.Add "Projects. Inactive", "Proyectos. Inactivos"


' -----------------------------------------------------------------------------
' Project registration

dictLabelFrench.Add "PROJECT REGISTRATION", "ENREGISTREMENT DE PROJET"
dictLabelSpanish.Add "PROJECT REGISTRATION", "REGISTRAR PROYECTO"

dictLabelFrench.Add "Project status", "Statut de projet"
dictLabelSpanish.Add "Project status", "Estatus del proyecto"

dictLabelFrench.Add "Country / Region", "Pays / Région"
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
dictLabelFrench.Add "If you have already registered your profile", "Si vous avez déjà enregistré votre profil, veuillez svp cliquer sur ce <a href=""login.asp" & AddUrlParams(sParams, "url=" + sScriptFullName) & """>lien pour mettre à jour votre CV</a>."
dictLabelSpanish.Add "If you have already registered your profile", "Si ya ha registrado su perfil, por favor <a href=""login.asp" & AddUrlParams(sParams, "url=" + sScriptFullName) & """> acceda para actualizar su CV</a>."

dictLabelEnglish.Add "Please fill in all the relevant information", "Please fill in all the relevant information and as many details on your experience as possible."
dictLabelFrench.Add "Please fill in all the relevant information", "Veuillez compléter toutes les informations importantes et autant de détails sur votre expérience que possible."
dictLabelSpanish.Add "Please fill in all the relevant information", "Por favor complete toda la informacion relevante y tantos detalles de su experiencia como sea posible."

dictLabelEnglish.Add "Fields marked with *", "Fields marked with <span class=""fcmp"">*</span> are required information."
dictLabelFrench.Add "Fields marked with *", "Les Champs identifiés par <span class=""fcmp"">*</span> sont obligatoires."
dictLabelSpanish.Add "Fields marked with *", "Los campos identicifacos con <span class=""fcmp"">*</span> requiren ser completados."

dictLabelEnglish.Add "You can always go back", "You can always go back and edit each section by clicking on the menu at the top."
dictLabelFrench.Add "You can always go back", "Vous pouvez toujours retourner en arrière et éditer chaque section en cliquant sur le menu au dessus."
dictLabelSpanish.Add "You can always go back", "Simepre puede retroceder y editar cada seccion pinchando en el menu situado arriba."

dictLabelFrench.Add "PERSONAL INFORMATION", "INFORMATIONS PERSONNELLES"
dictLabelSpanish.Add "PERSONAL INFORMATION", "INFORMACION PERSONAL"

dictLabelFrench.Add "CV language", "Langue du CV"
dictLabelSpanish.Add "CV language", "Lenguaje de CV"

dictLabelEnglish.Add "Personal title", "Title"
dictLabelFrench.Add "Personal title", "Civilité"
dictLabelSpanish.Add "Personal title", "Titulo"

dictLabelFrench.Add "Please select", "Choisissez svp"
dictLabelSpanish.Add "Please select", "Pofavor eliga"

dictLabelFrench.Add "First name", "Prénom(s)"
dictLabelSpanish.Add "First name", "Nombre"

dictLabelFrench.Add "Middle name", "Deuxième prénom"
dictLabelSpanish.Add "Middle name", "Segundo nombre"

dictLabelFrench.Add "Family name", "Nom"
dictLabelSpanish.Add "Family name", "Apellido"

dictLabelFrench.Add "Last name", "Nom"
dictLabelSpanish.Add "Last name", "Apellido"

dictLabelFrench.Add "Surname", "Nom"
dictLabelSpanish.Add "Surname", "Apellido"

dictLabelFrench.Add "Name", "Nom"
dictLabelSpanish.Add "Name", "Nombre"

dictLabelFrench.Add "Surname(s) / First name(s)", "Nom(s) / Prénom(s)"
dictLabelSpanish.Add "Surname(s) / First name(s)", "Apellido(s) / Nombre(s)"

dictLabelEnglish.Add "Date of birth", "Date of birth"
dictLabelFrench.Add "Date of birth", "Date de naissance"
dictLabelSpanish.Add "Date of birth", "Fecha de nacimiento"

dictLabelFrench.Add "Day", "jour"
dictLabelSpanish.Add "Day", "dia"

dictLabelFrench.Add "Month", "mois"
dictLabelSpanish.Add "Month", "mes"

dictLabelFrench.Add "Year", "année"
dictLabelSpanish.Add "Year", "año"

dictLabelEnglish.Add "Place of birth", "Place&nbsp;of&nbsp;birth"
dictLabelFrench.Add "Place of birth", "Lieu de naissance"
dictLabelSpanish.Add "Place of birth", "Lugar de naciemiento"

dictLabelFrench.Add "Civil status", "État civil"
dictLabelSpanish.Add "Civil status", "Estado civil"

dictLabelFrench.Add "Nationality", "Nationalité"
dictLabelSpanish.Add "Nationality", "Nacionalidad"

dictLabelEnglish.Add "Add nationality", "Click on <b>Add</b> to add a selected nationality to your list."
dictLabelFrench.Add "Add nationality", "Cliquez sur <b>ajouter</b> pour ajouter une nationalité choisie à votre liste."
dictLabelSpanish.Add "Add nationality", "Pinche en <b>ajouter</b> para incluir la nacionalidad elegido a su lista."

dictLabelEnglish.Add "Remove nationality", "If you want to remove a nationality, highlight it and click on <b>Remove</b>"
dictLabelFrench.Add "Remove nationality", "Si vous voulez supprimer une nationalité, sélectionnez-la et<br/>&nbsp; cliquez sur <b>supprimer</b>"
dictLabelSpanish.Add "Remove nationality", ""

dictLabelFrench.Add "Add", "Ajoutez"
dictLabelSpanish.Add "Add", "Añadir"

dictLabelFrench.Add "Remove", "Supprimer"
dictLabelSpanish.Add "Remove", "Suprimir"

dictLabelFrench.Add "Gender", "Sexe"
dictLabelSpanish.Add "Gender", "Sexo"

dictLabelFrench.Add "male", "masculin"
dictLabelSpanish.Add "male", "masculino"

dictLabelFrench.Add "female", "féminin"
dictLabelSpanish.Add "female", "femenino"

dictLabelFrench.Add "Marital status", "État civil"
dictLabelSpanish.Add "Marital status", "Estado civil"

dictLabelFrench.Add "Primary phone", "Téléphone"
dictLabelSpanish.Add "Primary phone", "Telefono principal"

dictLabelFrench.Add "Phone", "Téléphone"
dictLabelSpanish.Add "Phone", "Telefono"

dictLabelFrench.Add "Primary email", "Email"
dictLabelSpanish.Add "Primary email", "Email principal"

dictLabelFrench.Add "Status", "Statut"
dictLabelSpanish.Add "Status", "Estado"

dictLabelFrench.Add "Current position", "Poste actuel"
dictLabelSpanish.Add "Current position", "Posicion actual"

dictLabelFrench.Add "Present position", "Situation présente"
dictLabelSpanish.Add "Present position", "Situacion actual"

dictLabelFrench.Add "Desired employment / Occupational field", "Poste visé"
dictLabelSpanish.Add "Desired employment / Occupational field", "Empleo deseado / campo de ocupacion"

dictLabelFrench.Add "Key qualifications", "Principales compétences"
dictLabelSpanish.Add "Key qualifications", "Competencias principales"

dictLabelFrench.Add "Years of professional experience", "Nombre d'années d’expérience"
dictLabelSpanish.Add "Years of professional experience", "Numero de años de experiencia"

dictLabelFrench.Add "use only numbers", "utilisez uniquement des chiffres"
dictLabelSpanish.Add "use only numbers", "utilize solo numeros"

dictLabelFrench.Add "Specific experience in the region", "Expérience spécifique dans la région"
dictLabelSpanish.Add "Specific experience in the region", "Experiencia especifica en la region"

dictLabelFrench.Add "Specific experience in non industrialised countries", "Expérience spécifique dans la région"
dictLabelSpanish.Add "Specific experience in non industrialised countries", "Experiencia especifica en paises no industralizados"

dictLabelFrench.Add "SIRET number", "Numéro SIRET"
dictLabelSpanish.Add "SIRET number", "Numero SIRET"

' -----------------------------------------------------------------------------
' CV registration step 2 (register2.asp)

dictLabelFrench.Add "Education", "Éducation"
dictLabelSpanish.Add "Education", "Educacion"

dictLabelFrench.Add "EDUCATION", "ÉDUCATION"
dictLabelSpanish.Add "EDUCATION", "EDUCACION"

dictLabelFrench.Add "EDUCATION AND TRAINING", "ÉDUCATION ET FORMATION"
dictLabelSpanish.Add "EDUCATION AND TRAINING", "EDUCACION Y FORMACION"

dictLabelFrench.Add "Education and training", "Éducation et formation"
dictLabelSpanish.Add "Education and training", "Educacion y formacion"

dictLabelFrench.Add "Please fill in the institution name", "Veuillez compléter le institution"
dictLabelSpanish.Add "Please fill in the institution name", "Por favor rellene la institucion"

dictLabelFrench.Add "Please fill in the education end date", "Veuillez compléter la date de fin de la formation"
dictLabelSpanish.Add "Please fill in the education end date", "Por favor rellene las fechas del fin de la formacion"

dictLabelFrench.Add "Please fill in the education dates properly", "Veuillez insérer les dates de formation correctement"
dictLabelSpanish.Add "Please fill in the education dates properly", "Por favor incluya los datos de formacion correctas"

dictLabelFrench.Add "Please specify a type of diploma or degree obtained", "Veuillez spécifier un type de diplôme ou diplôme universitire obtenu"
dictLabelSpanish.Add "Please specify a type of diploma or degree obtained", "Por favor especifique el tipo de diploma o carrera universitaria obtenido"

dictLabelFrench.Add "Please specify the education subject", "Veuillez spécifier les matieres d'enseignement"
dictLabelSpanish.Add "Please specify the education subject", "Por favor rellene el campo de educacion"

dictLabelFrench.Add "No.", "Numéro"
dictLabelSpanish.Add "No.", "Numero"

dictLabelEnglish.Add "Institution name", "Institution&nbsp;name"
dictLabelFrench.Add "Institution name", "Institution&nbsp;où&nbsp;le<br />diplôme&nbsp;a&nbsp;été&nbsp;obtenu"
dictLabelSpanish.Add "Institution name", "Nombre&nbsp;de&nbsp;institucion"

dictLabelFrench.Add "Institution", "Institution"
dictLabelSpanish.Add "Institution", "Institucion"

dictLabelFrench.Add "Start date", "Date&nbsp;de&nbsp;début"
dictLabelSpanish.Add "Start date", "Fecha de comienzo"

dictLabelFrench.Add "Date from", "Date début"
dictLabelSpanish.Add "Date from", "Fecha de"

dictLabelFrench.Add "End date", "Date&nbsp;de&nbsp;fin"
dictLabelSpanish.Add "End date", "Fecha de fin"

dictLabelFrench.Add "Date to", "Date fin"
dictLabelSpanish.Add "Date to", "Fecha a"

dictLabelFrench.Add "Dates (from – to)", "Date (début - fin)"
dictLabelSpanish.Add "Dates (from – to)", "Fechas (comienzo - fin)"

dictLabelFrench.Add "Subject", "Sujet"
dictLabelSpanish.Add "Subject", "Campo"

dictLabelFrench.Add "Modify", "Modifiez"
dictLabelSpanish.Add "Modify", "Modificar"

dictLabelFrench.Add "Delete", "Supprimez"
dictLabelSpanish.Add "Delete", "Borrar"

dictLabelEnglish.Add "Type of diploma", "Type of Diploma /<br />Degree obtained"
dictLabelFrench.Add "Type of diploma", "Type de diplôme obtenu"
dictLabelSpanish.Add "Type of diploma", "Tipo de diploma obtenido"

dictLabelFrench.Add "Degree(s) or Diploma(s) obtained", "Diplôme(s) obtenu(s)"
dictLabelSpanish.Add "Degree(s) or Diploma(s) obtained", "Diplomas obtenidos"

dictLabelFrench.Add "If other please specify", "Si autre veuillez spécifier"
dictLabelSpanish.Add "If other please specify", "Si otro, por favor especifique"

dictLabelFrench.Add "If needed, please specify the exact title of your diploma", "Si nécessaire spécifiez le titre exact de votre diplôme"
dictLabelSpanish.Add "If needed, please specify the exact title of your diploma", "Por favor especifique el titulo exacto de su diploma"

dictLabelFrench.Add "If needed, please specify the exact title of your degree", "Si nécessaire spécifiez le titre exact de votre diplôme"
dictLabelSpanish.Add "If needed, please specify the exact title of your degree", "Por favor especifique el titulo exacto de su diploma"

dictLabelFrench.Add "Exact title of your degree", "Le titre exact de votre diplôme"
dictLabelSpanish.Add "Exact title of your degree", "El titulo exacto de su diploma"

dictLabelFrench.Add "Name and type of organisation providing education and training", "Nom et type de l'établissement d'enseignement ou de formation"
dictLabelSpanish.Add "Name and type of organisation providing education and training", "Nombre y tipo de establecimiento de formacion"

dictLabelFrench.Add "Principal subjects/occupational skills covered", "Principales matières/compétences professionnelles couvertes"
dictLabelSpanish.Add "Principal subjects/occupational skills covered", "Principales temas cubiertos"

dictLabelFrench.Add "Title of qualification awarded", "Intitulé du certificat ou diplôme délivré"
dictLabelSpanish.Add "Title of qualification awarded", "Titulo de diploma obtenido"

dictLabelFrench.Add "Level in national or international classification", "Niveau dans la classification nationale ou internationale"
dictLabelSpanish.Add "Level in national or international classification", "Nivel en la clasificacion nacional o internacional"


' -----------------------------------------------------------------------------
' CV registration step 2 (register21.asp)

dictLabelFrench.Add "Training", "Autre formation"
dictLabelSpanish.Add "Training", "Otra formacion"

dictLabelFrench.Add "TRAINING", "AUTRE FORMATION"
dictLabelSpanish.Add "TRAINING", "OTRA FORMACION"

dictLabelFrench.Add "Please fill in the training title", "Veuillez compléter le titre de formation"
dictLabelSpanish.Add "Please fill in the training title", "Por favor rellene el titulo de la formacion"

dictLabelFrench.Add "Please fill in the training end date", "Veuillez insérer la date de fin de formation"
dictLabelSpanish.Add "Please fill in the training end date", "Por favor rellene la fecha de fin de la formacion"

dictLabelFrench.Add "Please fill in the training dates properly", "Veuillez insérer les dates de formation correctement"
dictLabelSpanish.Add "Please fill in the training dates properly", "Por favor, agregue la informacion correcta"

dictLabelFrench.Add "Not specified", "Non spécifié"
dictLabelSpanish.Add "Not specified", "Sin especificar"

dictLabelFrench.Add "Title", "Titre"
dictLabelSpanish.Add "Title", "Titulo"

dictLabelFrench.Add "Skills / Qualifications", "Qualifications"
dictLabelSpanish.Add "Skills / Qualifications", "Cualificaciones"

dictLabelFrench.Add "Achievements", "Accomplissements"
dictLabelSpanish.Add "Achievements", "Logros"


' -----------------------------------------------------------------------------
' CV registration step 3 (register3.asp)

dictLabelFrench.Add "Professional experience", "Expériences professionnelles"

dictLabelFrench.Add "Please specify the project title or the name of the company or organisation", "Veuillez spécifier le titre du projet ou le nom de la compagnie ou de l'organisation"

dictLabelFrench.Add "Please fill in the experience start date", "Veuillez insérer la date de début d'expérience"

dictLabelFrench.Add "Please fill in the experience end date", "Veuillez insérer la date de fin d'expérience"

dictLabelFrench.Add "Please fill in the experience dates properly", "Veuillez insérer les dates d'expérience correctement"

dictLabelFrench.Add "Please fill in your position", "Veuillez compléter votre position"

dictLabelFrench.Add "Please make the description of the project shorter", "Veuillez rendre la description du projet plus courte"

dictLabelFrench.Add "Please select at least one country", "Veuillez choisir au moins un pays"

dictLabelFrench.Add "You cannot select more than 30 countries for one project", "Vous ne pouvez pas choisir plus de 30 pays pour un projet"

dictLabelFrench.Add "Please select at least one sub-sector of expertise", "Veuillez choisir au moins un sous-secteur d'expertise"

dictLabelFrench.Add "You cannot select more than 50 sectors for one project", "Vous ne pouvez pas choisir plus de 50 secteurs pour un projet"

dictLabelFrench.Add "Project title", "Titre du projet"

dictLabelEnglish.Add "Type of experience (Reg)", "Type of experience<br /><small>(if relevant)</small>"
dictLabelFrench.Add "Type of experience (Reg)", "Type d'expérience<br /><small>(le cas échéant)</small>"

dictLabelEnglish.Add "Project title (Reg)", "Project title<br /><small>(if relevant)</small>"
dictLabelFrench.Add "Project title (Reg)", "Titre du projet<br /><small>(le cas échéant)</small>"

dictLabelEnglish.Add "Main project features (Reg)", "Main project features<br /><small>(if relevant)</small>"
dictLabelFrench.Add "Main project features (Reg)", "Caractéristiques principales du projet<br /><small>(le cas échéant)</small>"

dictLabelFrench.Add "Position", "Position"

dictLabelFrench.Add "Project / Organisation", "Projet / organisation"

dictLabelFrench.Add "Company / Organisation", "Compagnie / organisation"

dictLabelFrench.Add "Position / Responsibility", "Position / responsabilité"

dictLabelFrench.Add "Beneficiary", "Bénéficiaire"

dictLabelFrench.Add "Location", "Lieu"

dictLabelFrench.Add "Countries", "Pays"

dictLabelFrench.Add "Sectors", "Secteurs"

dictLabelFrench.Add "Client references", "Références&nbsp;du&nbsp;client"

dictLabelFrench.Add "Company and reference person", "Société et personne de référence"

dictLabelEnglish.Add "Brief description of tasks", "Brief description of<br />the tasks assigned"
dictLabelFrench.Add "Brief description of tasks", "Courte description <br />des tâches assignées"

dictLabelEnglish.Add "Description of tasks", "Description of<br />the tasks assigned"
dictLabelFrench.Add "Description of tasks", "Description <br />des tâches assignées"

dictLabelFrench.Add "Funding agency", "Agence&nbsp;de&nbsp;placement"

dictLabelFrench.Add "Major funding agencies", "Principaux organismes de financement"

dictLabelFrench.Add "Other funding agencies", "Autres organismes de financement"

dictLabelEnglish.Add "Select funding agency from list", "Select funding agency from the list or specify in the field above if it is not in the list"
dictLabelFrench.Add "Select funding agency from list", "Choisissez l'agence de placement à partir de la liste ou spécifiez dans le domaine ci-dessus s'il n'est pas dans la liste"

dictLabelFrench.Add "SELECT PROJECT'S COUNTRIES", "CHOISISSEZ LES PAYS DU PROJET"

dictLabelFrench.Add "SELECT PROJECT'S SUB-SECTORS", "CHOISISSEZ LES SOUS-SECTEURS DU PROJET"

dictLabelEnglish.Add "KEY QUALIFICATION AND SPECIFIC EXPERIENCE", "KEY QUALIFICATION AND SPECIFIC EXPERIENCE (PROJECTS, ETC.)"
dictLabelFrench.Add "KEY QUALIFICATION AND SPECIFIC EXPERIENCE", "PRINCIPALES QUALIFICATIONS ET EXPERIENCES SPÉCIFIQUE (PROJETS, ETC)"

dictLabelEnglish.Add "Key qualification and specific experience", "Key qualification and specific experience (projects, etc.)"
dictLabelFrench.Add "Key qualification and specific experience", "Principales qualifications et experiences spécifique (projets, etc)"

dictLabelFrench.Add "PROFESSIONAL EXPERIENCE", "EXPÉRIENCES PROFESSIONNELLES"

dictLabelFrench.Add "EMPLOYMENT RECORD AND COMPLETED PROJECTS", "EMPLOIS PASSES ET PROJETS RÉALISÉS"
dictLabelFrench.Add "Employment record and completed projects", "Emplois passes et projets réalisés"

dictLabelFrench.Add "WORK EXPERIENCE", "EXPÉRIENCE PROFESSIONNELLE"

dictLabelFrench.Add "Work experience", "Expérience professionnelle"

dictLabelFrench.Add "Ongoing", "En cours"

dictLabelFrench.Add "Reference", "Référence"

dictLabelFrench.Add "contact person", "personne à contacter"

dictLabelFrench.Add "Occupation or position held", "Fonction ou poste occupé"

dictLabelFrench.Add "Main activities and responsibilities", "Principales activités et responsabilités"

dictLabelFrench.Add "Name and address of employer", "Nom et adresse de l'employeur"

dictLabelFrench.Add "Type of business or sector", "Type ou secteur d’activité"

dictLabelEnglish.Add "Specify professional experiences (GIP)", "Specify your main professional experiences in France or abroad.<br />Fill in the complete form for each experience, starting with the most recent one. Once you have filled in the form for one experience, click on [ Add an Experience ] to add the information related to your previous experiences."
dictLabelFrench.Add "Specify professional experiences (GIP)", "Spécifiez vos principales expériences en France ou ailleurs.<br />Remplissez le formulaire complet pour chaque expérience, commençant par la plus récente. Une fois vouz auriez rempli le formulaire pour une expérience, clicquer sur [ Ajouter une expérience ] pour ajouter l'information concernant vos expériences antérieures."

dictLabelEnglish.Add "SELECT PROJECT'S COUNTRIES (GIP)", "SELECT COUNTRIES IN WHICH <br/>&nbsp; &nbsp; &nbsp; &nbsp; YOU HAVE WORKED DURING THIS EXPERIENCE"
dictLabelFrench.Add "SELECT PROJECT'S COUNTRIES (GIP)", "CHOISISSEZ LE OU LES PAYS OU VOUS AVEZ <br/>&nbsp; &nbsp; &nbsp; &nbsp; ACCOMPLI CETTE EXPERIENCE/MISSION"

dictLabelFrench.Add "only for international projects", "uniquement pour les projets internationaux"

dictLabelEnglish.Add "First click on the sector title (GIP)", "First click on the sector title in the left column and sub-sectors will appear in the right column.<br />Then select the sub-sectors that best match your experience."
dictLabelFrench.Add "First click on the sector title (GIP)", "Cliquez d’abord sur le secteur, dans la colonne de gauche et les sous-secteurs apparaîtront dans la colonne de droite. Choisissez alors les sous-secteurs qui correspondent au mieux à l’expérience que vous décrivez. "

' -----------------------------------------------------------------------------
' CV registration step 4 (register4.asp)

dictLabelFrench.Add "Languages", "Langues"

dictLabelFrench.Add "Please select your native language", "S'il vous plaît sélectionnez votre langue maternelle"

dictLabelFrench.Add "Please choose a language", "S'il vous plaît choisissez une langue"

dictLabelFrench.Add "Please choose the levels of your knowledge", "S'il vous plaît choisissez le niveau de vos connaissances"

dictLabelFrench.Add "You are only allowed 20 languages", "Vous êtes seulement autorisé 20 langues"

dictLabelEnglish.Add "Add selected language", "Click on [ Add ] button to add a selected language to your list.<br />If you want to remove a language, highlight it and click on [ Remove ] button."
dictLabelFrench.Add "Add selected language", "Sélectionnez votre ou vos langues maternelles et cliques sur [ Ajouter ] pour les ajouter à la liste. Si vous souhaitez supprimer une langue, sélectionnez-la et cliquez sur [ Enlever ]."

dictLabelEnglish.Add "Choose a language and specify your level...", "Choose a language and specify your level (reading, speaking, writing). To add another languages click on [ Add language ] button and specify the levels for each of them."
dictLabelFrench.Add "Choose a language and specify your level...", "Choisissez une langue et spécifiez votre niveau (lu, parlé, écrit). Pour ajouter une langue cliquez sur [ Ajouter une langue ] et spécifiez à chaque fois les niveaux pour chacune d'entre elles."

dictLabelFrench.Add "Language", "Langue"

dictLabelFrench.Add "Native", "langue maternelle"

dictLabelFrench.Add "Mother tongue(s)", "Langue(s) maternelle(s)"

dictLabelFrench.Add "Other language(s)", "Autre(s) langue(s)"

dictLabelFrench.Add "Reading", "Lu"

dictLabelEnglish.Add "Reading(EP)", "Reading"
dictLabelFrench.Add "Reading(EP)", "Lire"

dictLabelFrench.Add "Understanding", "Comprendre"

dictLabelFrench.Add "Speaking", "Parlé"

dictLabelEnglish.Add "Speaking(EP)", "Speaking"
dictLabelFrench.Add "Speaking(EP)", "Parler"

dictLabelFrench.Add "Listening", "Écouter"

dictLabelFrench.Add "Writing", "Ecrit"

dictLabelEnglish.Add "Writing(EP)", "Writing"
dictLabelFrench.Add "Writing(EP)", "Écrire"

dictLabelFrench.Add "NATIVE LANGUAGES", "LANGUE MATERNELLES"
dictLabelFrench.Add "Native languages", "Langue maternelles"

dictLabelFrench.Add "FOREIGN LANGUAGES", "LANGUES ÉTRANGÈRES"
dictLabelFrench.Add "Foreign languages", "Langues étrangères"

dictLabelFrench.Add "LANGUAGES SKILLS", "CAPACITÉS LINGUISTIQUES"
dictLabelFrench.Add "Languages skills", "Capacités linguistiques"

dictLabelEnglish.Add "Languages skills EC", "Language skills: Indicate competence on a scale of 1 to 5 (1 - excellent; 5 - basic)"
dictLabelFrench.Add "Languages skills EC", "Connaissances linguistiques: Indiquer vos connaissances sur une échelle de 1 à 5 (1 - niveau excellent; 5 - niveau rudimentaire)"

dictLabelFrench.Add "Common European Framework of Reference for Languages", "Cadre européen commun de référence pour les langues"

dictLabelFrench.Add "Spoken interaction", "Prendre part à une conversation"

dictLabelFrench.Add "Spoken production", "S’exprimer oralement en continu"

' -----------------------------------------------------------------------------
' CV registration step 5 (register5.asp)

dictLabelEnglish.Add "Contact details & availability", "Contact&nbsp;details<br />&amp;&nbsp;availability"
dictLabelFrench.Add "Contact details & availability", "Coordonnées<br />&amp;&nbsp;disponibilité"

dictLabelFrench.Add "Contact details", "Coordonnées"

dictLabelFrench.Add "availability", "disponibilité"

dictLabelFrench.Add "permanent address", "adresse permanente"

dictLabelFrench.Add "Please fill in a street of ", "S'il vous plaît remplir la rue de "

dictLabelFrench.Add "Please fill in a city of ", "S'il vous plaît remplir la ville de "

dictLabelFrench.Add "Please fill in a postcode of ", "S'il vous plaît remplir le code postal de "

dictLabelFrench.Add "Please select a country of ", "S'il vous plaît sélectionner le pays de "

dictLabelFrench.Add "Please fill in ", "S'il vous plaît remplir "

dictLabelFrench.Add " permanent phone number.", " numéro de téléphone."

dictLabelFrench.Add "Please specify ", "S'il vous plaît précisez "

dictLabelFrench.Add " permanent email.", " adresse email."

dictLabelFrench.Add "Please retype ", "S'il vous plaît retaper "

dictLabelFrench.Add "permanent email correctly", "adresse email correctement"

dictLabelFrench.Add "current email correctly", "actuel courriel correctement"

dictLabelFrench.Add "Please make text of your availibility shorter.", "S'il vous plaît rendre le texte de votre disponibilité plus court."

dictLabelFrench.Add "Street", "Rue"

dictLabelFrench.Add "City", "Ville"

dictLabelFrench.Add "Postcode", "Code postal"

dictLabelFrench.Add "Country", "Pays"

dictLabelFrench.Add "Mobile", "GSM"

dictLabelFrench.Add "Fax", "Fax"

dictLabelFrench.Add "Website", "Site Web"

dictLabelEnglish.Add "Please specify availability", "Please specify the periods in which you are available in the next two years.<br />To guarantee the best matches, please keep your availability information updated."
dictLabelFrench.Add "Please specify availability", "Veuillez spécifier les périodes durant lesquelles vous êtes disponible pendant les deux années à venir. Merci de maintenir à jour les informations relatives à vos disponibilités."

dictLabelFrench.Add "Availability", "Disponibilité"

dictLabelFrench.Add "Availability & preferences", "Disponibilité & préférences"

dictLabelFrench.Add "location", "pays"

dictLabelEnglish.Add "Please state your preferences", "Please state your preferences for short-and/or long-term missions."
dictLabelFrench.Add "Please state your preferences", "Veuillez spécifier vos préférences pour des missions courtes et / ou de long terme."

dictLabelFrench.Add "Preferences", "Préférences"

dictLabelFrench.Add "Short-term missions", "Missions à court terme"

dictLabelFrench.Add "Long-term missions", "Missions à long terme"

dictLabelFrench.Add "Other skills", "Autres compétences"

dictLabelFrench.Add "Other preferences", "D'autres préférences"

dictLabelFrench.Add "Membership in professional bodies", "Adhésion à des associations ou corps professionnels"

dictLabelFrench.Add "Membership of professional bodies", "Affiliation à une organisation professionnelle"

dictLabelFrench.Add "Publications", "Publications"

dictLabelFrench.Add "References", "Références"

dictLabelFrench.Add "PERMANENT ADDRESS", "ADRESSE PERMANENTE"
dictLabelFrench.Add "Permanent address", "Adresse permanente"

dictLabelFrench.Add "CURRENT ADDRESS", "ADRESSE ACTUELLE"
dictLabelFrench.Add "Current address", "Adresse actuelle"

dictLabelFrench.Add "CURRENT ADDRESS (if different)", "ADRESSE ACTUELLE (SI ELLE EST DIFFÉRENTE)"
dictLabelFrench.Add "Current address (if different)", "Adresse actuelle (si elle est différente)"

dictLabelFrench.Add "CURRENT AVAILABILITY", "DISPONIBILITÉ ACTUELLE"
dictLabelFrench.Add "Current availability", "Disponibilité actuelle"

dictLabelFrench.Add "MISCELLANEOUS", "DIVERS"
dictLabelFrench.Add "Miscellaneous", "Divers"

dictLabelFrench.Add "ASSIGNMENT PREFERENCES", "PRÉFÉRENCE DES MISSIONS"
dictLabelFrench.Add "Assignment preferences", "Préférences des missions"

dictLabelFrench.Add "OTHER", "AUTRE"
dictLabelFrench.Add "Other", "Autre"

dictLabelFrench.Add "Address", "Adresse"

dictLabelFrench.Add "Other relevant information (e.g., Publications)", "Autres informations pertinentes (p,ex., références de publications)"

dictLabelFrench.Add "Personal skills and competences", "Aptitudes et compétences personnelles"

dictLabelFrench.Add "Social skills and competences", "Aptitudes et compétences sociales"

dictLabelFrench.Add "Organisational skills and competences", "Aptitudes et compétences organisationnelles"

dictLabelFrench.Add "Technical skills and competences", "Aptitudes et compétences techniques"

dictLabelFrench.Add "Computer skills and competences", "Aptitudes et compétences informatiques"

dictLabelFrench.Add "Artistic skills and competences", "Aptitudes et compétences artistiques"

dictLabelFrench.Add "Other skills and competences", "Autres aptitudes et compétences"

dictLabelFrench.Add "Driving licence(s)", "Permis de conduire"

dictLabelFrench.Add "Additional information", "Information complémentaire"

' -----------------------------------------------------------------------------
' CV WB format 

dictLabelFrench.Add "FORM TECH-6", "FORMULAIRE TECH-6"
dictLabelFrench.Add "CURRICULUM VITAE (CV) FOR PROPOSED PROFESSIONAL STAFF", "CURRICULUM VITAE (CV) DU PERSONNEL CLE PROPOSE"

dictLabelFrench.Add "Proposed Position", "Poste"

dictLabelEnglish.Add "only one candidate...", "only one candidate shall be nominated for each position"
dictLabelFrench.Add "only one candidate...", "un seul candidat par poste"

dictLabelFrench.Add "Name of Firm", "Nom du consultant"

dictLabelFrench.Add "insert name of firm proposing the staff", "indiquer le nom de la société proposant le personnel"

dictLabelFrench.Add "Name of Staff", "Nom de l’employé"

dictLabelFrench.Add "Date of Birth", "Date de naissance"

dictLabelEnglish.Add "Education (WB)", "Education"
dictLabelFrench.Add "Education (WB)", "Formation"

dictLabelFrench.Add "Name of assignment or project", "Nom du projet ou de la mission"

dictLabelFrench.Add "Main project features", "Principales caractéristiques du projet"

dictLabelEnglish.Add "Indicate college/university...", "Indicate college/university and other specialized education of staff member, giving names of institutions, degrees obtained, and dates of obtainment"
dictLabelFrench.Add "Indicate college/university...", "Indiquer les études universitaires et autres études spécialisées de l’employé ainsi que les noms des institutions fréquentées, les diplômes obtenus et les dates auxquelles ils l’ont été"

dictLabelFrench.Add "Membership of Professional Associations", "Affiliation à des associations/groupements professionnels"

dictLabelFrench.Add "Membership in Professional Societies", "Affiliation à des associations professionnelles"

dictLabelFrench.Add "Other Training", "Autres formations"

dictLabelEnglish.Add "Indicate significant training...", "Indicate significant training since degrees under 5 - Education were obtained"
dictLabelFrench.Add "Indicate significant training...", "Indiquer toute autre formation reçue depuis 5 ci-dessus"

dictLabelFrench.Add "Countries of Work Experience", "Pays où l’employé a travaillé"

dictLabelFrench.Add "List countries where staff has worked in the last ten years", "Donner la liste des pays ou l’employé a travaillé au cours des 10 dernières années"

dictLabelEnglish.Add "For each language indicate proficiency...", "For each language indicate proficiency: good, fair, or poor in speaking, reading, and writing"
dictLabelFrench.Add "For each language indicate proficiency...", "Indiquer pour chacune le degré de connaissance : bon, moyen, médiocre pour ce qui est de la langue parlée, lue et écrite"

dictLabelFrench.Add "Employment Record", "Expérience professionnelle"

dictLabelEnglish.Add "Starting with present position...", "Starting with present position, list in reverse order every employment held by staff member since graduation, giving for each employment (see format here below): dates of employment, name of employing organization, positions held"
dictLabelFrench.Add "Starting with present position...", "En commençant par son poste actuel, donner la liste par ordre chronologique inverse de tous les emplois exercés par l’employé depuis la fin de ses études. Pour chaque emploi (voir le formulaire ci-dessous), donner les dates, le nom de l’employeur et le poste occupé."

'dictLabelFrench.Add "", "Depuis [année]"

'dictLabelFrench.Add "", "jusqu’à [année]"

dictLabelFrench.Add "Employer", "Employeur"

dictLabelFrench.Add "Position held", "Poste"

dictLabelFrench.Add "Activities performed", "Activités réalisées"

dictLabelFrench.Add "Description of duties", "Description des tâches assignées"

dictLabelFrench.Add "Detailed Tasks Assigned", "Détail des tâches exécutées"

dictLabelFrench.Add "Detailed Tasks Assigned (AFB)", "Attributions spécifiques"

dictLabelFrench.Add "List all tasks to be performed under this assignment", "Indiquer toutes les tâches à exécuter dans le cadre de cette proposition"

dictLabelFrench.Add "Work Undertaken that Best Illustrates Capability to Handle the Tasks Assigned", "Expérience de l’employé qui illustre le mieux sa compétence"

dictLabelEnglish.Add "Among the assignments in which the staffs have been involved...", "Among the assignments in which the staffs have been involved, indicate the following information for those assignments that best illustrate staff capability to handle the tasks listed under point 11."
dictLabelFrench.Add "Among the assignments in which the staffs have been involved...", "Donner notamment les informations suivantes qui illustrent au mieux la compétence professionnelle de l’employé pour les tâches mentionnées au point 11."

dictLabelEnglish.Add "Among the assignments in which the staffs have been involved (ADB)...", "Among the assignments in which the staffs have been involved, indicate the following information for those assignments that best illustrate staff capability to handle the tasks listed under point 12."
dictLabelFrench.Add "Among the assignments in which the staffs have been involved (ADB)...", "Donner notamment les informations suivantes qui illustrent au mieux la compétence professionnelle de l’employé pour les tâches mentionnées au point 12."

dictLabelFrench.Add "Certification", "Attestation"

dictLabelEnglish.Add "I, the undersigned, certify...", "I, the undersigned, certify that to the best of my knowledge and belief, this CV correctly describes me, my qualifications, and my experience.  I understand that any wilful misstatement described herein may lead to my disqualification or dismissal, if engaged."
dictLabelFrench.Add "I, the undersigned, certify...", "Je, soussigné, certifie, en toute conscience, que les renseignements ci-dessus rendent fidèlement compte de ma situation, de mes qualifications et de mon expérience. J’accepte que toute déclaration volontairement erronée peut entraîner mon exclusion, ou mon renvoi si j’ai été engagé."

dictLabelEnglish.Add "I, the undersigned, certify (AFB)...", "I, the undersigned, certify that to the best of my knowledge and belief, these biodata correctly describe myself, my qualifications and my experience."
dictLabelFrench.Add "I, the undersigned, certify (AFB)...", "Je, soussigné, certifie, sur la base des données à ma disposition, que les renseignements ci-dessus rendent fidèlement compte de ma situation, de mes qualifications et de mon expérience."

dictLabelFrench.Add "Signature of staff member or authorized representative of the staff", "Signature de l’employé et du représentant habilité du consultant"

dictLabelFrench.Add "Day/Month/Year", "Jour/mois/année"

dictLabelFrench.Add "Full name of authorized representative", "Nom du représentant habilité"

dictLabelFrench.Add "Years with Firm", "Années d’emploi au sein de la firme"

dictLabelEnglish.Add "Key Qualifications (AFD)", "Key Qualifications"
dictLabelFrench.Add "Key Qualifications (AFD)", "Principales qualifications"
%>