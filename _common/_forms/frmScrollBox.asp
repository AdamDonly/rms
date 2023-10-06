<% 
Dim iCouScrllHeight, iDonScrllHeight, iSctScrllHeight
Dim full_shift, is_mozilla, is_nav, is_chrome, is_ie_new
Dim bShowAllDonors
Dim iMaxSectors
' For experts' SIP and member's Free Trial BSC - limitation on 80 sectors 
' SIP updates from IBF's network are enabled for unlimited number of sectors
If (sScriptFileName="sc_register.asp") Or (sScriptFileName="sc_update.asp" And Left(sUserIpAddress, 10)<>"158.29.157") Or (sScriptFileName="bsc_register.asp" And (sAccessType="trial" Or sAccessType="cml")) Then
	iMaxSectors=80
Else
	iMaxSectors=420
End If

' to identify later
Dim ts, tt, k1, k2, mreg, mExF, m, j, mExT, mExTInt, strTemp, cp, mGZ, mNt, mNtInt, mOrg, mOrgInt

If aNt<12 Then
	iCouScrllHeight=60+21*aNt
Else
	iCouScrllHeight=311
End If

If aOrg<=9 Then
	iDonScrllHeight=26+21*aOrg
Else
	iDonScrllHeight=224+3
End If

If aExT<12 Then
	iSctScrllHeight=60+21*aExT+full_shift*13
Else
	iSctScrllHeight=396
End If


sUserAgent=LCase(Request.ServerVariables("HTTP_USER_AGENT"))
If (InStr(sUserAgent,"mozilla")>0) AND (InStr(sUserAgent,"spoofer")=0) AND (InStr(sUserAgent,"compatible")=0) AND (InStr(sUserAgent,"opera555")=0) AND (InStr(sUserAgent,"webtv")=0) AND (InStr(sUserAgent,"gecko")=0) Then
	is_nav = 1
Else
	is_nav = 0
End If
If (InStr(sUserAgent,"mozilla")>0) AND (InStr(sUserAgent,"gecko")>0) Then
	is_mozilla = 1
Else
	is_mozilla = 0
End If
If (InStr(sUserAgent,"chrome")>0) Then
	is_chrome = 1
	is_mozilla = 0
Else
	is_chrome = 0
End If
If (InStr(sUserAgent,"msie 10")>0 Or InStr(sUserAgent,"msie 9")>0) Then
	is_ie_new = 1
Else
	is_ie_new = 0
End If

Sub InsertScrollStyles
	If bShowAllDonors=0 Then
		iDonScrllHeight=iDonScrllHeight-24
	End If	
	%>
	<style>
	#Div1 {position:relative; width: 391; height: <%=iCouScrllHeight%>; left: 0; top:0; clip:rect(0, 391, <%=iCouScrllHeight%>, 0); overflow:hidden; visibility:hidden; }
	#DivText1 {position:absolute; top:0; left:0} 

	#Div2 {position:relative; width: 391; height: <%=iDonScrllHeight%>; left: 0; top:0; clip:rect(0, 391, <%=iDonScrllHeight%>, 0); overflow:hidden; visibility:hidden; }
	#DivText2 {position:absolute; top:0; left:0} 

	#Div3 {position:relative; width: 391; height: <%=iSctScrllHeight%>; left: 0; top:0; clip:rect(0, 391, <%=iSctScrllHeight%>, 0); overflow:hidden; visibility:hidden; }
	#DivText3 {position:absolute; top:0; left:0} 

	</style>
<%
End Sub


Sub InsertJSScrollFunctions(bShowTotal, bShowAll)
%>
<script language="JavaScript">
<!--
jNtInt = new Array();
jNtCode= new Array();
jNtZone= new Array();
jGZnInt = new Array();
jOrgInt = new Array();
jOrgCode = new Array();
jOrgMain = new Array();
jMDonInt = new Array();
jExTInt= new Array();
jExTCode= new Array();
jExTSrch= new Array();
jExFInt = new Array();

var rst_cou=new Array();
var rst_don=new Array();
var rst_sct=new Array();
var agt;

var cntCou=0;
var cntDon=0;
var cntSct=0;

Chk1 = new Image(13,13); Chk1.src = 'n.gif';
Chk2 = new Image(13,13); Chk2.src = 'c.gif';

// **********************************************
function RoundTo5(amount)
{
var addamount, newamount, lastdigit;
newamount=Math.round(amount);
newamountstring=''+newamount
lastdigit=newamountstring.substr(newamountstring.length-1, 1);
if (lastdigit==0) {addamount=0;}
if (lastdigit==1) {addamount=4;}
if (lastdigit==2) {addamount=3;}
if (lastdigit==3) {addamount=2;}
if (lastdigit==4) {addamount=1;}
if (lastdigit==5) {addamount=0;}
if (lastdigit==6) {addamount=4;}
if (lastdigit==7) {addamount=3;}
if (lastdigit==8) {addamount=2;}
if (lastdigit==9) {addamount=1;}

return(newamount+addamount);
}

// **********************************************
function GetControl(layerName){
  if(document.getElementById) return document.getElementById(layerName)
  if(document.all) return document.all[layerName]
  if(document.layers) return eval('document.' + layerName)
}

// **********************************************
function ChangeControlColor(lControl, lItemNum, lColorNum)
{
var tColor, tControl;
  tControl=GetControl(lControl + 'Text' + lItemNum.toString());
  if(lColorNum==1) {tColor='#CC0000';} else {tColor='#000066';}
  tControl.style.color=tColor;
  //tControl.color=tColor;
}

// **********************************************
function ChangeControlText(lControl, lItemNum, lItemValue )
{
var tControl, tPos;
if (is_ie4up || is_mozilla) 
{
  <% If is_nav<>1 Then %>
  tControl=GetControl(lControl + 'Text' + lItemNum.toString());
  tPos=tControl.innerHTML.indexOf('(');
  if(lItemValue>0)
	{if(tPos>0)
	{tControl.innerHTML=tControl.innerHTML.substring(0,tPos-1);}
	tControl.innerHTML=tControl.innerHTML + ' (' + lItemValue.toString() + ')';
	  if(lItemValue>0)
	  {ChangeControlColor(lControl, lItemNum, 1);}
	}
  else
	{tControl.innerHTML=tControl.innerHTML.substring(0,tPos-1);
	ChangeControlColor(lControl, lItemNum, 0);}
  <% End If %>
}
}


// **********************************************
function RCou(num)
{
var dm;
dm = (is_nav4up) ? document.Div1.document.DivText1.document : document;

        if (num>0)
		{if (jNtInt[num]==0)
		{ mmb_cou++ ; jNtInt[num]=jNtCode[num]; dm.images['couInt'+num].src = Chk2.src;
		jGZnInt[jNtZone[num]]++;
		ChangeControlText('Reg', jNtZone[num]-1, jGZnInt[jNtZone[num]]);
		}
		else
		{ mmb_cou-- ; jNtInt[num]=0; dm.images['couInt'+num].src = Chk1.src;
		  dm.images['regInt'+jNtZone[num]].src = Chk1.src; jGZnInt[jNtZone[num]]--;
		<% If bShowAll=1 Then %>
		  dm.images['couIntA'].src = Chk1.src; jNtIntAll=0; 
		<% End If %>
		ChangeControlText('Reg', jNtZone[num]-1, jGZnInt[jNtZone[num]]);
		}}
	<% If bShowAll=1 Then %>
	else
		{if (jNtIntAll==0)
	        { jNtIntAll=1; mmb_cou=<%=aNt%>; dm.images['couIntA'].src = Chk2.src;
		  <% For i=0 To aGZ-1 %>jGZnInt[<%=i+1%>]=<%=aGZnScroll(i)%>; <% Next %>
		  for (j=1; j<=<%=aGZ%>; j++)
		  {dm.images['regInt'+j].src = Chk2.src;
		   ChangeControlText('Reg', j-1, jGZnInt[j]);
		  }
		  for (j=1; j<=<%=aNt%>; j++)
		  {jNtInt[j]=jNtCode[j]; dm.images['couInt'+j].src = Chk2.src;}
		}
	  	else
		{ jNtIntAll=0; mmb_cou=0; dm.images['couIntA'].src = Chk1.src;
		  <% For i=0 To aGZ-1 %>jGZnInt[<%=i+1%>]=0; <% Next %>
		  for (j=1; j<=<%=aGZ%>; j++)
		  {jGZnInt[j]=0; dm.images['regInt'+j].src = Chk1.src;
		   ChangeControlText('Reg', j-1, 0);
 		  }
		  for (j=1; j<=<%=aNt%>; j++)
		  {jNtInt[j]=0; dm.images['couInt'+j].src = Chk1.src;}
		}}
	<% End If %>
<% If bShowTotal Then %>
document.RegFormCou.mmb_cou_total.value=mmb_cou;
SetTotal();
<% End If %>
}


// **********************************************
function RReg(num)
{
var dm;
var jGZnTmp = new Array();
dm = (is_nav4up) ? document.Div1.document.DivText1.document : document;

  		if (jGZnInt[num]==0)
		{
		for (j=1; j<=<%=aNt%>; j++)
			{ if (jNtZone[j]==num && jNtInt[j]==0)
			{ dm.images['couInt'+j].src = Chk2.src;
			jNtInt[j]=jNtCode[j]; mmb_cou++;}}
		<% For i=0 To aGZ-1 %>jGZnTmp[<%=i+1%>]=<%=aGZnScroll(i)%>; <% Next %>
		jGZnInt[num]=jGZnTmp[num]; dm.images['regInt'+num].src = Chk2.src;
		ChangeControlText('Reg', num-1, jGZnInt[num]);
		}
		else
		{
		for (j=1; j<=<%=aNt%>; j++)
			{ if (jNtZone[j]==num && jNtInt[j]!=0)
			{ dm.images['couInt'+j].src = Chk1.src;
			jNtInt[j]=0; mmb_cou-- ; 
			<% If bShowAll=1 Then %>
			  dm.images['couIntA'].src = Chk1.src; jNtIntAll=0; 
			<% End If %>
			}}
		jGZnInt[num]=0; dm.images['regInt'+num].src = Chk1.src; 
		ChangeControlText('Reg', num-1, 0);
		}
<% If bShowTotal Then %>
document.RegFormCou.mmb_cou_total.value=mmb_cou;
SetTotal();
<% End If %>
}


function RDon(num)
{
var dm;
dm = (is_nav4up) ? document.Div2.document.DivText2.document : document;

        if (num>0)
		{if (jOrgInt[num]==0)
		{ mmb_don++; jOrgInt[num]=jOrgCode[num]; dm.images['donInt'+num].src = Chk2.src;
		jMDonInt[jOrgMain[num]+2]++;
		ChangeControlText('MDon', jOrgMain[num]+1, jMDonInt[jOrgMain[num]+2]);
		}
		else
		{ mmb_don--; jOrgInt[num]=0; dm.images['donInt'+num].src = Chk1.src;
		<% If bShowAll=1 Then %>
		  jOrgIntAll=0; dm.images['donIntA'].src = Chk1.src;
		<% End If %>
		jMDonInt[jOrgMain[num]+2]--;
		ChangeControlText('MDon', jOrgMain[num]+1, jMDonInt[jOrgMain[num]+2]);
		}}
	<% If bShowAll=1 Then %>
	else
		{if (jOrgIntAll==0)
	        { mmb_don=<%=aOrg%>; jOrgIntAll=1; dm.images['donIntA'].src = Chk2.src;
		  for (j=1; j<=<%=aOrg%>; j++)
		  {jOrgInt[j]=jOrgCode[j]; dm.images['donInt'+j].src = Chk2.src;}
		  jMDonInt[1]=9; jMDonInt[2]=17; 
		  for (j=1; j<=2; j++)
		  {ChangeControlText('MDon', j-1, jMDonInt[j]);}
		}
	  	else
		{ mmb_don=0; jOrgIntAll=0; dm.images['donIntA'].src = Chk1.src;
		  for (j=1; j<=<%=aOrg%>; j++)
		  {jOrgInt[j]=0; dm.images['donInt'+j].src = Chk1.src;}
		  jMDonInt[1]=0; jMDonInt[2]=0;
		  for (j=1; j<=2; j++)
		  {ChangeControlText('MDon', j-1, jMDonInt[j]);}
		}}
	<% End If %>

<% If bShowTotal Then %>
document.RegFormDon.mmb_don_total.value=mmb_don;
SetTotal();
<% End If %>
}

// **********************************************
function RSct(num)
{
var dm;
dm = (is_nav4up) ? document.Div3.document.DivText3.document : document;

if (mmb_sct>=<%=iMaxSectors%> && jExTInt[num]==0)
{ alert('You can not select more than <%=iMaxSectors%> sectors!');}
else
{
	// Checking total number of Fields of Interest
        if (num>0)
		{if (jExTInt[num]==0)
		{ mmb_sct++ ; jExTInt[num]=jExTCode[num]; dm.images["sctInt"+num].src = Chk2.src;
		jExFInt[jExTSrch[num]]++;
		ChangeControlText('MSct', jExTSrch[num]-1, jExFInt[jExTSrch[num]]);
		}
		else
		{ mmb_sct-- ; jExTInt[num]=0; dm.images["sctInt"+num].src = Chk1.src;
		<% If bShowAll=1 Then %>
		  dm.images["sctIntA"].src = Chk1.src; jExTIntAll=0;
		<% End If %>
		  dm.images["msctInt"+jExTSrch[num]].src = Chk1.src; 
		jExFInt[jExTSrch[num]]--;
		ChangeControlText('MSct', jExTSrch[num]-1, jExFInt[jExTSrch[num]]);
		}}
	<% If bShowAll=1 Then %>
	else
		{if (jExTIntAll==0)
	        { jExTIntAll=1; mmb_sct=<%=aExT%>; dm.images["sctIntA"].src = Chk2.src;
		  for (j=1; j<=<%=aExT%>; j++)
		  {jExTInt[j]=jExTCode[j]; dm.images["sctInt"+j].src = Chk2.src;}
		  <% For i=0 To aExF-1 %>jExFInt[<%=i+1%>]=<%=aExFScroll(i)%>; <% Next %>
		  for (j=1; j<=<%=aExF%>; j++)
		  {dm.images["msctInt"+j].src = Chk2.src;
		   ChangeControlText('MSct', j-1, jExFInt[j]);
		  }
		}
	  	else
		{ jExTIntAll=0; mmb_sct=0; dm.images["sctIntA"].src = Chk1.src;
		  for (j=1; j<=<%=aExT%>; j++)
		  {jExTInt[j]=0; dm.images["sctInt"+j].src = Chk1.src;}
		  <% For i=0 To aExF-1 %>jExFInt[<%=i+1%>]=0; <% Next %>
		  for (j=1; j<=<%=aExF%>; j++)
		  {dm.images["msctInt"+j].src = Chk1.src;
		   ChangeControlText('MSct', j-1, jExFInt[j]);
		  }
		  }}
	<% End If %>
}
<% If bShowTotal Then %>
document.RegFormSct.mmb_sct_total.value=mmb_sct;
SetTotal();
<% End If %>
}

// **********************************************
function RMsct(num)
{
var jExFTmp = new Array();
var dm; var j; var nselected=0;
dm = (is_nav4up) ? document.Div3.document.DivText3.document : document;

if (mmb_sct>=<%=iMaxSectors%> && jExFInt[num]==0)
{ alert('You can not select more than <%=iMaxSectors%> sectors!'); return;}
else
{
  		if (jExFInt[num]==0)
		{
		for (j=1; j<=<%=aExT%>; j++)
			{ if (jExTSrch[j]==num && jExTInt[j]==0)
			{ 
			if (mmb_sct>=<%=iMaxSectors%>)
			{ alert('You can not select more than <%=iMaxSectors%> sectors!'); jExFInt[num]=nselected; ChangeControlText('MSct', num-1, jExFInt[num]); return;}
			dm.images['sctInt'+j].src = Chk2.src;
			jExTInt[j]=jExTCode[j]; mmb_sct++; nselected++;}}
		<% For i=0 To aExF-1 %>jExFTmp[<%=i+1%>]=<%=aExFScroll(i)%>; <% Next %>
		jExFInt[num]=jExFTmp[num]; dm.images['msctInt'+num].src = Chk2.src;
		ChangeControlText('MSct', num-1, jExFInt[num]);
		}
		else
		{
		for (j=1; j<=<%=aExT%>; j++)
			{ if (jExTSrch[j]==num && jExTInt[j]!=0)
			{ dm.images['sctInt'+j].src = Chk1.src;
			jExTInt[j]=0; mmb_sct-- ;}}
		jExFInt[num]=0; dm.images['msctInt'+num].src = Chk1.src; 
		ChangeControlText('MSct', num-1, 0);
		<% If bShowAll=1 Then %>
		  dm.images["sctIntA"].src = Chk1.src; jExTIntAll=0;
		<% End If %>
		}
}
<% If bShowTotal Then %>
document.RegFormSct.mmb_sct_total.value=mmb_sct;
SetTotal();
<% End If %>
}


// **********************************************
<% If bShowTotal Then %>
function SetTotal(cfield)
{                 
<% If sScriptFileName<>"update_prf2.asp" And sScriptFileName<>"register_prf2.asp" Then %>
var cntTotal = document.RegForm.mmb_total_price.value;
var activeCurrency;
var activeExchangeRate;
activeExchangeRate=<%=Replace(iExchangeRate, ",", ".")%>;

  document.RegFormCou.mmb_cou_total.value=mmb_cou;
  document.RegFormDon.mmb_don_total.value=mmb_don;
  document.RegFormSct.mmb_sct_total.value=mmb_sct;

    <% If sAccessType<>"trial"  Or sAccessType<>"cml" Then %>
   if (cfield=='cur2')
   { document.RegForm.mmb_total_currency1.selectedIndex=document.RegForm.mmb_total_currency2.selectedIndex; }
   else if (cfield=='cur1')
   { document.RegForm.mmb_total_currency2.selectedIndex=document.RegForm.mmb_total_currency1.selectedIndex; }
    <% End If %>
  
  if (mmb_sct>0 && mmb_cou>0 && mmb_don>0)
  {
    if ((mmb_sct>0) && (mmb_sct<21))  {cntSct=1;}
    if ((mmb_sct>20) && (mmb_sct<41)) {cntSct=2;}
    if ((mmb_sct>40) && (mmb_sct<81)) {cntSct=3;}
    if ((mmb_sct>80) && (mmb_sct<161)){cntSct=4;}
    if (mmb_sct>160) { cntSct=5 }

    if ((mmb_cou>0) && (mmb_cou<16))  {cntCou=1;}
    if ((mmb_cou>15) && (mmb_cou<41)) {cntCou=2;}
    if ((mmb_cou>40) && (mmb_cou<81)) {cntCou=3;}
    if ((mmb_cou>80) && (mmb_cou<121)){cntCou=4;}
    if (mmb_cou>120) { cntCou=5 }

    cntDon=mmb_don;
    // ignoring additional funding agencies
    for (i=1; i<=<%=aOrg%>; i++)
    { if (jOrgInt[i]>0 && jOrgMain[i]==0) {--cntDon;}}

    if ((cntDon>0) && (cntDon<3)) {cntDon=1;}
    if ((cntDon>2) && (cntDon<5)) {cntDon=2;}
    if ((cntDon>4) && (cntDon<7)) {cntDon=3;}
    if ((cntDon>6) && (cntDon<20)) {cntDon=3;}

  <% If Not iExpertID>0 Then %>
    <% If sAccessType="trial" Or sAccessType="cml" Then %>
    <% Else %>
    for (var i=0;i<4;i++){
	if (document.RegForm.mmb_period[i].checked == true) 
	{document.RegForm.mmb_period_hid.value=i+1;}}

    // 90 eur - added subscription fee for devbusiness
    cntTotal= jPrice[cntSct][cntCou][cntDon];

    if (cntTotal>0)
    {activeCurrency=(document.RegForm.mmb_total_currency1.options[document.RegForm.mmb_total_currency1.selectedIndex].value=="EUR")?1:activeExchangeRate;
     document.RegForm.mmb_total_price.value=RoundTo5(cntTotal*activeCurrency);
     document.RegForm.mmb_total_sum.value=document.RegForm.mmb_total_price.value*document.RegForm.mmb_period_hid.value;}
    <% End If %>

  <% End If %>  
  }
  else
  {document.RegForm.mmb_total_price.value='';
  document.RegForm.mmb_total_sum.value='';}
<% End If %>
}
<% End If %>


function RestoreInt()
{
var dm;
var jGZnTmp = new Array();
var jMOrgTmp = new Array();
var jExFTmp = new Array();
if (document.RegForm.mmb_cou_hid.value.length>5)
{jNtInt = document.RegForm.mmb_cou_hid.value.split(',');
if (jNtInt[0]>'')
{mmb_cou=jNtInt[0];jNtInt[0]='';
dm = (is_nav4up) ? document.Div1.document.DivText1.document : document;
for (i=1;i<jNtInt.length;i++) {if (jNtInt[i]>0) {dm.images['couInt'+i].src = Chk2.src;	jGZnInt[jNtZone[i]]=jGZnInt[jNtZone[i]]+1;}}
<% For i=0 To aGZ-1 %>jGZnTmp[<%=i+1%>]=<%=aGZnScroll(i)%>; <% Next %>

for (i=1;i<jGZnInt.length;i++) {if (jGZnInt[i]>0) {ChangeControlText('Reg', i-1, jGZnInt[i]);} if (jGZnInt[i]==jGZnTmp[i]) {dm.images['regInt'+i].src = Chk2.src;} }
}}

if (document.RegForm.mmb_don_hid.value.length>5)
{jOrgInt = document.RegForm.mmb_don_hid.value.split(',');
if (jOrgInt[0]>'')
{mmb_don=jOrgInt[0];jOrgInt[0]='';
dm = (is_nav4up) ? document.Div2.document.DivText2.document : document;
for (i=1;i<jOrgInt.length;i++) {if (jOrgInt[i]>0) {dm.images['donInt'+i].src = Chk2.src; jMDonInt[2+jOrgMain[i]]=jMDonInt[2+jOrgMain[i]]+1;}}
for (i=1;i<jMDonInt.length;i++) {if (jMDonInt[i]>0) {ChangeControlText('MDon', i-1, jMDonInt[i]);}}
}}

if (document.RegForm.mmb_sct_hid.value.length>5)
{jExTInt = document.RegForm.mmb_sct_hid.value.split(',');
if (jExTInt[0]>'')
{mmb_sct=jExTInt[0];jExTInt[0]='';
dm = (is_nav4up) ? document.Div3.document.DivText3.document : document;
for (i=1;i<jExTInt.length;i++) {if (jExTInt[i]>0) {dm.images['sctInt'+i].src = Chk2.src; jExFInt[jExTSrch[i]]=jExFInt[jExTSrch[i]]+1;}}
<% For i=0 To aExF-1 %>jExFTmp[<%=i+1%>]=<%=aExFScroll(i)%>; <% Next %>
for (i=1;i<jExFInt.length;i++) {if (jExFInt[i]>0) {ChangeControlText('MSct', i-1, jExFInt[i]);} if (jExFInt[i]==jExFTmp[i]) {dm.images['msctInt'+i].src = Chk2.src;} }
}}
<% If bShowTotal Then %>
SetTotal('');
<% End If %>
}

function LoadInt()
{

if (document.RegForm.mmb_cou_hid.value.length>5)
{jNtInt = document.RegForm.mmb_cou_hid.value.split(',');
if (jNtInt[0]>'')
{mmb_cou=jNtInt[0];jNtInt[0]='';
<% If bShowTotal Then %>
document.RegFormCou.mmb_cou_total.value=mmb_cou;
<% End If %>
}}

if (document.RegForm.mmb_don_hid.value.length>5)
{jOrgInt = document.RegForm.mmb_don_hid.value.split(',');
if (jOrgInt[0]>'')
{mmb_don=jOrgInt[0];jOrgInt[0]='';
<% If bShowTotal Then %>
document.RegFormDon.mmb_don_total.value=mmb_don;
<% End If %>
}}

if (document.RegForm.mmb_sct_hid.value.length>5)
{jExTInt = document.RegForm.mmb_sct_hid.value.split(',');
if (jExTInt[0]>'')
{mmb_sct=jExTInt[0];jExTInt[0]='';
<% If bShowTotal Then %>
document.RegFormSct.mmb_sct_total.value=mmb_sct;
<% End If %>
}}

<% If bShowTotal Then %>
SetTotal();
<% End If %>
}

// **********************************************
function ShowPrice(sSender) 
{  
  var params, pfile;
  if (sSender=='bsc_reg')
  {
	if (document.RegForm.mmb_total_currency1.options[document.RegForm.mmb_total_currency1.selectedIndex].value=="USD")
	{pfile='bsc_price.asp?dcr=USD'}
	else
	{pfile='bsc_price.asp?dcr=EUR'}
	params=(cntCou>0 && cntDon>0 && cntSct>0)? '?prm=1&cou='+cntCou+'&don='+cntDon+'&sct='+cntSct : '';
  }
  window.open(pfile+params,'ANWnd','scrollbars=yes,status=yes,resizable=yes,menubar=yes');
}

var Sct_ScrollList_Height = 0;
if (document.getElementById('DivText3')) {
	Sct_ScrollList_Height=document.getElementById('DivText3').clientHeight;
}

// -->
</script>
<%
End Sub


Sub ShowCouScrollBox(sBoxTitle, sLeftMenuTitle, bShowLeftMenu, bShowAll, bShowRegions, bShowCountries, bShowTotal)
%>
	<% If sBoxTitle>"" Then %>
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br>
		<table width="100%" cellpadding=0 cellspacing=0 border=0>
		<tr><td width="50%"><p class="fttl"><img src="<%=sHomePath%>image/<%=imgFormBullet%>" width=7 height=7 align=left vspace=3 hspace=8><%=sBoxTitle%></p></td>
		<td width="50%" align="right"><p class="txt"><a class="tmenu" href="<%=sHomePath%>_data/cou_lst.rtf" target=_blank>Download&nbsp;the&nbsp;list&nbsp;of&nbsp;countries</a>&nbsp;&nbsp;&nbsp;&nbsp;</p></td>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
	<% End If %>
	
	<table  cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td colspan=5 width=578 bgcolor="<%=colFormHeaderTop%>" valign="top"><img src="<% =sHomePath %>image/x.gif" width=498 height=1 vspace=0><br>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td bgcolor="<%=colFormHeaderTop%>" width=169 valign="top">
		<% Set objTempRs=objConn.Execute("EXEC usp_DatContinentSelect")
		ts=0
		tt=24
		If bShowAll=0 Then
			tt=0
		End If

		If sLeftMenuTitle>"" Then
			Response.Write("<p class=""ftxt"">" & sLeftMenuTitle & "</p>")
		Else
			Response.Write("<img src=""../image/x.gif"" width=120 height=5><br>")
		End If

		j=0
		While Not objTempRs.Eof
		FOR i=0 to aGZ-1 
		If aGZnContinent(i)=objTempRs("id_Continent") Then 
		If i>0 Then
		  ts=ts+aGZnScroll(i-1)*21+34+tt-is_nav*(aGZnScroll(i-1)*2+3+tt/6)-is_chrome*(aGZnScroll(i-1)*1.07)+is_mozilla*(aGZnScroll(i-1)*2.3) + is_ie_new*(aGZnScroll(i-1)*0.9)

		  tt=0
		End If %>
		<p class="fsml"><img src="x.gif" width=3 height=7 vspace=2 hspace=4 align="left"><a class="dk" name="Reg<%=i%>" id="Reg<%=i%>" href="javascript:;" onClick="scroll(7000,1);scroll(-7000,1);scroll(<%=ts%>,1);noScroll();"><span name="RegText<%=i%>" id="RegText<%=i%>"><%=aGZnInfo(i)%></span></a></p><% End If %><% NEXT %>
		<% objTempRs.MoveNext
		Response.Write("<img src='x.gif' width=120 height=6><br>")
		j=j+1
		WEnd 
		objTempRs.Close
		Set objTempRs=Nothing 
		%>
	</td>
	<td bgcolor="<%=colFormBodyRight%>" width=1><img src="<% =sHomePath %>image/x.gif" width=1 height=50></td>
	<td bgcolor="<%=colFormBodyText%>" width=392 valign=top>
		<script language="Javascript">
		if (is_ie4up || is_nav4up){document.writeln('<DIV id=Div1><DIV id=DivText1>');}
		</script>
		<% If bShowAll=1 Then %>
			<p class="<%=cssScrllMainTitle%>"><a href="javascript:RCou(0);"><img src="n.gif" name='couIntA' vspace=2 hspace=10 border=0 align="left"></a>Select all countries</p>
		<% End If %>
		<% k1=0
		k2=0
		FOR i=0 to aGZ-1 
		k1=k1+1%>
		<% mreg=0 %>
		<% FOR m=0 to mGZ-1 %><% If (aGZnCode(i)=mGZnCode(m)) AND (aGZnScroll(i)=mGZnScroll(m)) Then %><% mreg=mGZnCode(m) %><% Response.Write("<script language=JavaScript>jGZnInt["& m+1 &"]=1;</script>") %><% Exit For %><% End If %><% NEXT %>
			<p class="<%=cssScrllSubTitle%>"><a href="javascript:RReg(<%=k1%>);"><img src='<% If mreg>0 Then %>c.gif<% Else %>n.gif<% End If %>' name='regInt<%=k1%>' vspace=2 hspace=10 border=0 align=left></a><%=aGZnInfo(i)%><br>[ select all countries in the region ]</p>
		<% FOR j=0 to aNt-1 %>
		<% If aNtZone(j)=aGZnCode(i) Then 
		k2=k2+1 %>
		<% mreg=0 %>
		<% FOR m=0 to mNt-1 %><% If aNtCode(j)=mNtCode(m) Then %><% mreg=mNtCode(m) %><% Exit For %><% End If %><% NEXT %>
		<% mNtInt = mNtInt &","& mreg %>
			<p class="<%=cssScrllText%>"><a href="javascript:RCou(<%=k2%>);"><img src='<% If mreg>0 Then %>c.gif<% Else %>n.gif<% End If %>' name='couInt<%=k2%>' align=left border=0></a><%=aNtInfo(j)%></p><% End If %>
		<% NEXT %><% NEXT %>
			<p class="fslh"><img src="x.gif" width=391 height=1></p>

		<script language="Javascript">
		if (is_ie4up || is_nav4up) {document.writeln('</DIV></DIV>');}
		</script>
		</td>
		<td bgcolor="FFFFFF" width=1><img src="x.gif" width=1 height=50></td>
		<td width=15 bgcolor="FFFFFF" valign="top" background="<% =sHomePath %>image/vn_scrl.gif">
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-3220,1);noScroll();"><img src="<% =sHomePath %>image/vn_uup.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-66,1);noScroll();" onmouseover="scroll(-8,1);"><img src="<% =sHomePath %>image/vn_up.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-66,1);noScroll();"><img src="<% =sHomePath %>image/x.gif" width=15 height=<%=Int(iCouScrllHeight/2)-30%> border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(66,1);noScroll();"><img src="<% =sHomePath %>image/x.gif" width=15 height=<%=iCouScrllHeight-Int(iCouScrllHeight/2)-30%> border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(66,1);noScroll();" onmouseover="scroll(8,1);"><img src="<% =sHomePath %>image/vn_dn.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(3220,1);noScroll();"><img src="<% =sHomePath %>image/vn_ddn.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">

	<% ' Total items selected text box
	If bShowTotal=1 Then 
	%>
	<tr height=1><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table border=0 cellspacing=0 cellpadding=0 width="100%">
		<tr><form method="post" name="RegFormCou">
		<td width=170 valign="top"><p class="ftxt">Total selected:</td>
		<td width=408><input type="text" name="mmb_cou_total" readOnly maxLength=3 size=6 onBlur="SetTotal('chk',0)"></td>
		</form></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<% 
	End If 
	%>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
<%
End Sub


Sub ShowDonScrollBox(sBoxTitle, sLeftMenuTitle, bShowLeftMenu, bShowAll, bShowRegions, bShowCountries, bShowTotal, bDonorsDelimited)
%>

	<% If sBoxTitle>"" Then %>
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br><p class="fttl"><img src="<%=sHomePath%>image/<%=imgFormBullet%>" width=7 height=7 align=left vspace=3 hspace=8><%=sBoxTitle%></p></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
	<% End If %>
	
	<table  cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td colspan=5 width=578 bgcolor="<%=colFormHeaderTop%>" valign="top"><img src="<% =sHomePath %>image/x.gif" width=498 height=1 vspace=0><br>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td bgcolor="<%=colFormHeaderTop%>" width=169 valign="top">
		<% If sLeftMenuTitle>"" Then
			Response.Write("<p class=""fsml""><img src=""x.gif"" width=3 height=100 vspace=2 hspace=4 align=""left"">" & sLeftMenuTitle & "</p>")
		Else
			Response.Write("<img src=""../image/x.gif"" width=120 height=5><br>")
		End If
		If bShowLeftMenu>0 Then %>
		<p class="fsml"><img src="x.gif" width=3 height=7 vspace=2 hspace=4 align="left"><a class="dk" name="MDon0" id="MDon0" href="javascript:;" onClick="scroll(9000,2);scroll(-9000,2);scroll(0,2);noScroll();"><span name="MDonText0" id="MDonText0">Major funding agencies</span></a></p>
		<p class="fsml"><img src="x.gif" width=3 height=7 vspace=2 hspace=4 align="left"><a class="dk" name="MDon1" id="MDon1" href="javascript:;" onClick="scroll(9000,2);scroll(-9000,2);scroll(228,2);noScroll();"><span name="MDonText1" id="MDonText1">Bonus funding agencies</span></a></p>
		<% Else %>
		<img src="x.gif" width=1 height=7><br>		
		<p class="fsml"><a class="dk" name="MDon0" id="MDon0" href="javascript:;" onClick="scroll(9000,2);scroll(-9000,2);scroll(0,2);noScroll();"><span name="MDonText0" id="MDonText0">Major funding agencies</span></a></p>
		<p class="fsml"><a class="dk" name="MDon1" id="MDon1" href="javascript:;" onClick="scroll(9000,2);scroll(-9000,2);scroll(190,2);noScroll();"><span name="MDonText1" id="MDonText1">Other funding agencies</span></a></p>
		<% End If %>
	</td>
	<td bgcolor="<%=colFormBodyRight%>" width=1><img src="<% =sHomePath %>image/x.gif" width=1 height=50></td>
	<td bgcolor="<%=colFormBodyText%>" width=392 valign=top>
		<script language="Javascript">
		if (is_ie4up || is_nav4up ){ document.writeln('<DIV id=Div2><DIV id=DivText2>');}
		</script>
		<% If bShowAll=1 Then %>
			<p class="<%=cssScrllMainTitle%>"><a href="javascript:RDon(0);"><img src="n.gif" name='donIntA' vspace=2 hspace=10 border=0 align="left"></a>Select all funding agencies</p>
		<% End If %>
		<% FOR j=0 to aOrg-1 %>
		  <% If aOrgMainDonor(j)=0 And bDonorsDelimited=0 Then 
			bDonorsDelimited=1
			If sScriptFileName="register_bsc.asp" Or sScriptFileName="update_bsc.asp" Then %>
				<p class="<%=cssScrllSubTitle%>">&nbsp;&nbsp;&nbsp;Free bonus funding agencies</p>
			<% End If %>
			<p class="fslb"><img src="x.gif" width=391 height=1></p>
			<p class="fslh"><img src="x.gif" width=391 height=1></p>

		  <% End If %>
		  <% mreg=0 %>
		  <% FOR m=0 to mOrg-1 %><% If aOrgCode(j)=mOrgCode(m) Then %><% mreg=mOrgCode(m) %><% Exit For %><% End If %><% NEXT %>
		  <% mOrgInt = mOrgInt &","& mreg %>		  
			<p class="<%=cssScrllText%>"><a href="javascript:RDon(<%=j+1%>);"><img src='<% If mreg>0 Then %>c.gif<% Else %>n.gif<% End If %>' name='donInt<%=j+1%>' align="left" border=0></a><%=aOrgInfo(j)%></p><% NEXT %>
			<p class="fslh"><img src="x.gif" width=391 height=1></p>
		<script language="Javascript">
		if (is_ie4up || is_nav4up){ document.writeln('</DIV></DIV>');}
		</script>
		</td>
		<td bgcolor="FFFFFF" width=1><img src="x.gif" width=1 height=50></td>
		<td width=15 bgcolor="FFFFFF" valign="top" background="<% =sHomePath %>image/vn_scrl.gif">
		<a href="javascript:;" onmouseout="noScroll();" onmouseover="scroll(-8,2);"><img src="<% =sHomePath %>image/vn_up.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-66,2);noScroll();"><img src="<% =sHomePath %>image/x.gif" width=15 height=<%=Int(iDonScrllHeight/2)-15%> border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(66,2);noScroll();"><img src="<% =sHomePath %>image/x.gif" width=15 height=<%=iDonScrllHeight-Int(iDonScrllHeight/2)-15%> border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onmouseover="scroll(8,2);"><img src="<% =sHomePath %>image/vn_dn.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">

	<% ' Total items selected text box
	If bShowTotal=1 Then 
	%>
	<tr height=1><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>

	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table border=0 cellspacing=0 cellpadding=0 width="100%">
		<tr><form method="post" name="RegFormDon">
		<td width=170 valign="top"><p class="ftxt">Total selected:</td>
		<td width=408><input type="text" name="mmb_don_total" readOnly maxLength=3 size=6 onBlur="SetTotal('chk',0)"></td>
		</form></tr>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<% 
	End If 
	%>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
<%
End Sub


Sub ShowSctScrollBox(sBoxTitle,  sLeftMenuTitle, bShowLeftMenu, bShowAll, bShowRegions, bShowCountries, bShowTotal)
%>
	<% If sBoxTitle>"" Then %>
	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=18><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1><br>
		<table width="100%" cellpadding=0 cellspacing=0 border=0>
		<tr><td width="50%"><p class="fttl"><img src="<%=sHomePath%>image/<%=imgFormBullet%>" width=7 height=7 align=left vspace=3 hspace=8><%=sBoxTitle%></p></td>
		<td width="50%" align="right"><p class="txt"><a class="tmenu" href="<%=sHomePath%>_data/sct_lst.rtf" target=_blank>Download&nbsp;the&nbsp;list&nbsp;of&nbsp;sectors</a>&nbsp;&nbsp;&nbsp;&nbsp;</p></td>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=18></td></tr>
	<tr height=1><td width=1 bgcolor="<%=colFormHeaderTop%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormHeaderBottom%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
	<% End If %>

	<table  cellspacing=0 cellpadding=0 border=0 align="center">
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td colspan=5 width=578 bgcolor="<%=colFormHeaderTop%>" valign="top"><img src="<% =sHomePath %>image/x.gif" width=498 height=1 vspace=0><br>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="x.gif" width=1 height=1></td>
	<td bgcolor="<%=colFormHeaderTop%>" width=169 valign="top">
		<% 
		If sLeftMenuTitle>"" Then
			Response.Write("<p class=""ftxt"">" & sLeftMenuTitle & "</p>")
		Else
			Response.Write("<img src=""../image/x.gif"" width=120 height=5><br>")
		End If

		ts=0
		tt=24
		If bShowAll=0 Then
			tt=0
		End If

		FOR i=0 to aExF-1
		If i>0 Then
		  ts=ts+aExFScroll(i-1)*21+34+tt+aExFShift(i-1)*13-is_chrome*(aExFScroll(i-1)*1+aExFShift(i-1)*0.35)+is_mozilla*(aExFScroll(i-1)*2.2+aExFShift(i-1)*2.15) + is_ie_new*(aExFScroll(i-1)*0.95)
		  tt=0
		End If %>

		<p class="fsml"><img src="x.gif" width=3 height=7 vspace=2 hspace=4 align="left"><a class="dk" name="MSct<%=i%>" id="MSct<%=i%>" href="javascript:;" onClick="scroll(12000,3);scroll(-12000,3);scroll(<%=ts%>,3);noScroll();"><span name="MSctText<%=i%>" id="MSctText<%=i%>"><%=aExFShort(i)%></span></a></p><% NEXT %>
		<img src="x.gif" width=7 height=4><br>
	</td>
	<td bgcolor="<%=colFormBodyRight%>" width=1><img src="x.gif" width=1 height=50></td>
	<td bgcolor="<%=colFormBodyText%>" width=392 valign=top>
		<script language="Javascript">
		if (is_ie4up || is_nav4up ){ document.writeln('<DIV id=Div3><DIV id=DivText3>');}</script>
		<% If bShowAll=1 Then %>
			<p class="<%=cssScrllMainTitle%>"><a href="javascript:RSct(0);"><img src="n.gif" name='sctIntA' vspace=2 hspace=10 border=0 align="left"></a>Select all sectors and sub-sectors</p>
		<% End If %>
		<% k1=0
		k2=0
		FOR i=0 to aExF-1
		k1=k1+1%>
		<% mreg=0 %>
		<% FOR m=0 to mExF-1 %><% If (aExFCode(i)=mExFCode(m)) AND (aExFScroll(i)=mExFScroll(m)) Then %><% mreg=mExFCode(m) %><% Response.Write("<script language=JavaScript>jExFInt["& m+1 &"]=1;</script>") %><% Exit For %><% End If %><% NEXT %>
			<p class="<%=cssScrllSubTitle%>"><a href="javascript:RMsct(<%=k1%>)"><img src='<% If mreg>0 Then %>c.gif<% Else %>n.gif<% End If %>' name='msctInt<%=k1%>' vspace=2 hspace=10 border=0 align="left"></a><%=aExFInfo(i)%><% If Len(aExFInfo(i))<54 Then%><br><% Else %> &nbsp; &nbsp; <% End If %>[ select all sub-sectors ]</p>
		<% FOR j=0 to aExT-1 %>
		<% If aExTSrch(j)=aExFCode(i) Then 
		 k2=k2+1 %>
		<% mreg=0 %>
		<% FOR m=0 to mExT-1 %><% If aExTCode(j)=mExTCode(m) Then %><% mreg=mExTCode(m) %><% Exit For %><% End If %><% NEXT %>
		<% mExTInt = mExTInt &","& mreg %>		  
			<p class="<%=cssScrllText%>"><a href="javascript:RSct(<%=k2%>);"><img src='<% If mreg>0 Then %>c.gif<% Else %>n.gif<% End If %>' name='sctInt<%=k2%>' align=left border=0></a><%=CutString(aExTInfo(j),54)%></p><% End If %>
		<% NEXT %><% NEXT %>
			<p class="fslh"><img src="x.gif" width=391 height=1></p>

		<script language="Javascript">
		if (is_ie4up || is_nav4up)
		{ document.writeln('</DIV></DIV>');}
		</script>

		</td>
		<td bgcolor="FFFFFF" width=1><img src="x.gif" width=1 height=50></td>
		<td width=15 bgcolor="FFFFFF" valign="top" background="<% =sHomePath %>image/vn_scrl.gif">
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-6000,3);noScroll();"><img src="<% =sHomePath %>image/vn_uup.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-66,3);noScroll();" onmouseover="scroll(-8,3);"><img src="<% =sHomePath %>image/vn_up.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(-66,3);noScroll();"><img src="<% =sHomePath %>image/x.gif" width=15 height=<%=Int(iSctScrllHeight/2)-30%> border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(66,3);noScroll();"><img src="<% =sHomePath %>image/x.gif" width=15 height=<%=iSctScrllHeight-Int(iSctScrllHeight/2)-30%> border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(66,3);noScroll();" onmouseover="scroll(8,3);"><img src="<% =sHomePath %>image/vn_dn.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
		<a href="javascript:;" onmouseout="noScroll();" onclick="scroll(6000,3);noScroll();"><img src="<% =sHomePath %>image/vn_ddn.gif" width=15 height=15 border=0 vspace=0 hspace=0></a><br>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>

	<table width=580 cellspacing=0 cellpadding=0 border=0 align="center">
	<% If bShowTotal Then %>
	<tr height=1><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	
	<tr><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderTop%>" valign="top">
		<table border=0 cellspacing=0 cellpadding=0 width="100%">
                <form method="post" name="RegFormSct">
		<tr>
		<td width=170 valign="top"><p class="ftxt">Total selected:</td>
		<td width=408><input type="text" name="mmb_sct_total" readOnly maxLength=3 size=6 onBlur="SetTotal('chk',0)"></td>
		</tr></form>
		</table>
	</td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<% End If %>

	<tr height=1><td width=1 bgcolor="<%=colFormBodyLeft%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="<%=colFormHeaderMain%>"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="<%=colFormBodyRight%>"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	<tr height=1><td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td>
	<td width=578 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=498 height=1></td>
	<td width=1 bgcolor="#FFFFFF"><img src="<% =sHomePath %>image/x.gif" width=1 height=1></td></tr>
	</table>
<%
End Sub
%>