<%
Const cCvTypeDisabled = 0
Const cCvTypeEnabled = 1

Dim bCvTypeActive
bCvTypeActive = cCvTypeEnabled

' List of the CV types enabled on the CVIP
Dim dictCvType
Set dictCvType = CreateObject("Scripting.Dictionary")

dictCvType.Add "Seureca", "Seureca"
dictCvType.Add "Veolia Group", "Veolia Group"
dictCvType.Add "Other Company", "Other Company"
dictCvType.Add "Independent", "Independent"
%>