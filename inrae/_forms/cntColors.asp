<%

'--------------------------------------------------------------------
'
' Color constants
'
'--------------------------------------------------------------------

' --- Form's colors
Dim colFormHeaderMain
Dim colFormHeaderTop
Dim colFormHeaderBottom
Dim colFormHeaderSplitter

Dim colFormBodyMenu 
Dim colFormBodyText 
Dim colFormBodyLeft 
Dim colFormBodyRight 
Dim colFormBodyBottom 
Dim imgFormBullet

Dim cssScrllMainTitle
Dim cssScrllSubTitle
Dim cssScrllText

If sColorScheme="RED" Then
	' Red scheme
	colFormHeaderMain = "#CC0000"
	colFormHeaderTop = "#FFD1BE"
	colFormHeaderBottom = "#861C00"
	colFormHeaderSplitter = "#FFFFFF"
	
	colFormBodyMenu = "#FFD1BE"
	colFormBodyText = "#FFF0E9"
	colFormBodyLeft = "#FFE2E2"
	colFormBodyRight = "#F36A58"
	colFormBodyBottom = "#CC0000"
	imgFormBullet="bbox.gif"

	cssScrllMainTitle = "frl1"
	cssScrllSubTitle = "frl2"
	cssScrllText = "frl3"
ElseIf sColorScheme="BLUE" Then
	' Blue scheme
	colFormHeaderMain = "#3366CC"
	colFormHeaderTop = "#97CAFB"
	colFormHeaderBottom = "#153B80"
	colFormHeaderSplitter = "#FFFFFF"
	
	colFormBodyMenu = "#97CAFB"
	colFormBodyText = "#E0F3FF"
	colFormBodyLeft = "#D1E7F7"
	colFormBodyRight = "#4694E1"
	colFormBodyBottom = "#3366CC"
	imgFormBullet="rbox.gif"

	cssScrllMainTitle = "fsl1"
	cssScrllSubTitle = "fsl2"
	cssScrllText = "fsl3"
ElseIf sColorScheme="TEAL" Then
	colFormHeaderMain = "#00847f"
	colFormHeaderTop = "#99cccc"
	colFormHeaderBottom = "#153B80"
	colFormHeaderSplitter = "#FFFFFF"
	
	colFormBodyMenu = "#99cccc"
	colFormBodyText = "#e6edec"
	colFormBodyLeft = "#D1E7F7"
	colFormBodyRight = "#009999"
	colFormBodyBottom = "#00847f"
	imgFormBullet="rbox.gif"

	cssScrllMainTitle = "fsl1"
	cssScrllSubTitle = "fsl2"
	cssScrllText = "fsl3"
End If

%>