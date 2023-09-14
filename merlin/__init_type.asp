<%
Const cCvTypeDisabled = 0
Const cCvTypeEnabled = 1

Dim bCvTypeActive
bCvTypeActive = cCvTypeDisabled

' List of the CV types enabled on the CVIP
Dim dictCvType
Set dictCvType = CreateObject("Scripting.Dictionary")
%>