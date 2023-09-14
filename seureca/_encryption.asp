<%
Const aExtraSalt = "SeUR3c@1422%521$Cv1p8fr"

Class CEncryption
	Private Sub Class_Initialize()
		Md5_Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
	End Sub

	Public Function Md5Encode(sText)
		Md5Encode = MD5(sText)
	End Function

	Public Function Md5EncodeWithSalt(sText)
		Md5EncodeWithSalt = Md5Encode(sText & aExtraSalt)
	End Function

	Public Function Md5DoubleEncodeWithSalt(sText1, sText2)
		Md5DoubleEncodeWithSalt = Md5Encode(Md5Encode(Md5Encode(sText1) & sText2) & aExtraSalt)
	End Function

	Public Function Sha1Encode(sText)
		Sha1Encode = hex_sha1(sText)
	End Function

	Public Function Sha1EncodeWithSalt(sText)
		Sha1EncodeWithSalt = Sha1Encode(sText & aExtraSalt)
	End Function

	Public Function Sha1DoubleEncodeWithSalt(sText1, sText2)
		Sha1DoubleEncodeWithSalt = Sha1Encode(Sha1Encode(Sha1Encode(sText1) & sText2) & aExtraSalt)
	End Function

	Public Function USha1Encode(sText)
		USha1Encode = UCase(Sha1Encode(sText))
	End Function

	%>
	<!-- #include file = "_security/sha1_js.asp" -->
	<!-- #include file = "_security/md5.asp" -->
	<%

	Function Sha1DotNetEncode(sText)
		Dim objUTF8Encoding, objCryptoServiceProvider
		Dim Bytes

		'Borrow some objects from .NET (supported from 1.1 onwards)
		Set objUTF8Encoding = CreateObject("System.Text.UTF8Encoding")
		Set objCryptoServiceProvider = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")

		'Convert the string to a byte array and hash it
		Bytes = objUTF8Encoding.GetBytes_4(sPassword & sSalt)
		Bytes = objCryptoServiceProvider.ComputeHash_2((Bytes))
		 
		Sha1DotNetEncode = Base64Encode(Stream_BinaryToString(Bytes))
	End Function

	Function Base64Encode(sText)
	    Dim oXML, oNode

	    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
	    Set oNode = oXML.CreateElement("base64")
	    oNode.dataType = "bin.base64"
	    oNode.nodeTypedValue = Stream_StringToBinary(sText)
	    Base64Encode = oNode.text
	    Set oNode = Nothing
	    Set oXML = Nothing
	End Function

	Function Base64Decode(ByVal vCode)
	    Dim oXML, oNode

	    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
	    Set oNode = oXML.CreateElement("base64")
	    oNode.dataType = "bin.base64"
	    oNode.text = vCode
	    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
	    Set oNode = Nothing
	    Set oXML = Nothing
	End Function

	'Stream_StringToBinary Function
	Function Stream_StringToBinary(Text)
		Const adTypeText = 2
		Const adTypeBinary = 1

		'Create Stream object
		Dim BinaryStream 'As New Stream
		Set BinaryStream = CreateObject("ADODB.Stream")

		'Specify stream type - we want To save text/string data.
		BinaryStream.Type = adTypeText

		'Specify charset For the source text (unicode) data.
		BinaryStream.CharSet = "us-ascii"

		'Open the stream And write text/string data To the object
		BinaryStream.Open
		BinaryStream.WriteText Text

		'Change stream type To binary
		BinaryStream.Position = 0
		BinaryStream.Type = adTypeBinary

		'Ignore first two bytes - sign of
		BinaryStream.Position = 0

		'Open the stream And get binary data from the object
		Stream_StringToBinary = BinaryStream.Read

		Set BinaryStream = Nothing
	End Function

	'Stream_BinaryToString Function
	Function Stream_BinaryToString(Binary)
		Const adTypeText = 2
		Const adTypeBinary = 1

		'Create Stream object
		Dim BinaryStream 'As New Stream
		Set BinaryStream = CreateObject("ADODB.Stream")

		'Specify stream type - we want To save binary data.
		BinaryStream.Type = adTypeBinary

		'Open the stream And write binary data To the object
		BinaryStream.Open
		BinaryStream.Write Binary

		'Change stream type To text/string
		BinaryStream.Position = 0
		BinaryStream.Type = adTypeText

		'Specify charset For the output text (unicode) data.
		BinaryStream.CharSet = "us-ascii"

		'Open the stream And get text/string data from the object
		Stream_BinaryToString = BinaryStream.ReadText
		Set BinaryStream = Nothing
	End Function

End Class

Dim Encryption
Set Encryption = New CEncryption
%>
