function CheckVAT(Lng) {

var CountryArr = new Array();

CountryArr[1]=548
CountryArr[2]=550
CountryArr[3]=551
CountryArr[4]=554
CountryArr[5]=555
CountryArr[6]=558
CountryArr[7]=559
CountryArr[8]=560
CountryArr[9]=563
CountryArr[10]=564
CountryArr[11]=565
CountryArr[12]=570
CountryArr[13]=569
CountryArr[14]=571
CountryArr[15]=574
//////////////
var c=0,d=0,e=0,f=0;
var opt=document.orgForm.orgCou.selectedIndex;
var lnk=0;


	for(i=1;i<=15;i++){if (document.orgForm.orgCou.options[opt].value == CountryArr[i]){lnk=1;}}
	///////////////////

   //if((document.orgForm.orgCou.options[opt].value == 550) &&(document.orgForm.Belgian.checked ==false))    {c=1;}
   //if((document.orgForm.orgCou.options[opt].value == 550) &&(document.orgForm.Belgian.checked ==true))    {d=1; }
  // if((document.orgForm.orgCou.options[opt].value != 550) &&(document.orgForm.Belgian.checked ==false))     { e=1;}
  // if((document.orgForm.orgCou.options[opt].value != 550) &&(document.orgForm.Belgian.checked ==true))     { f=1;}

  //if(lnk==1) {c=1;}
  //if(lnk==0) {d=1;}
  
  
   //if((document.orgForm.orgCou.options[opt].value == 550) &&(document.orgForm.Belgian.checked ==true))    {d=1; }
  // if((document.orgForm.orgCou.options[opt].value != 550) &&(document.orgForm.Belgian.checked ==false))     { e=1;}
  // if((document.orgForm.orgCou.options[opt].value != 550) &&(document.orgForm.Belgian.checked ==true))     { f=1;}


///////////////////
if(Lng=='Eng'){
/////////////////////
  

			//////case 1
			if(document.orgForm.orgCou.options[opt].value == 550){document.orgForm.VATCheck.value = "YES";}
			//////case 2			
			if((document.orgForm.orgCou.options[opt].value != 550)&&(lnk==1)){
			if(document.orgForm.orgVat.value == ""){ 
			if(confirm("Please fill the VAT number. Click on Cancel to fill VAT number. Or Click on OK if you do not have a VAT number."))
			{document.orgForm.VATCheck.value = "YES";}else{return false;}
		    }
		    else{document.orgForm.VATCheck.value = "NO";}
		    }
		    //////case 3
			if(lnk==0){	document.orgForm.VATCheck.value = "NO";}
		    
  
//if (c==1) {
//			alert("You are subjected to Belgian VAT.");
//			document.orgForm.VATCheck.value = "YES";
//			if(document.orgForm.orgVat.value == ""){alert("Please Fill in your VAT number."); document.orgForm.orgVat.select(); return false;}
//	      }
//if (d==1) {
//			document.orgForm.VATCheck.value = "YES";
//			if(document.orgForm.orgVat.value == ""){alert("Please Fill in your VAT number."); document.orgForm.orgVat.select(); return false;}			
//	      }
//if (e==1) {
//         if(document.orgForm.orgVat.value == ""){document.orgForm.VATCheck.value = "YES";}
//          else{document.orgForm.VATCheck.value = "NO";}
//          }	
	
//if (f==1) {	document.orgForm.VATCheck.value = "YES";	
  //      }	


////////::::to check the format of vat number according to Renoud
	if (document.orgForm.orgVat.value != ""){
	
	////////////new check for EU country
	if (lnk==1){
	//////////:
	
	if(document.orgForm.orgVat.value.length<8) {alert ("Please Fill correct VAT number. e.g 'BE417827795'"); document.orgForm.orgVat.select(); return false;}
	
	var alphabets;
	var digits;
	var VATInput;

	alphabets = "ABCDEFGHIJKLMNOPWRSTUVWXYZ0123456789-";

///	alphabets = "ABCDEFGHIJKLMNOPWRSTUVWXYZ";
//	digits="0123456789-";

	VATInput=document.orgForm.orgVat.value;
	for (i = 0; i < VATInput.length; i++){


			if (alphabets.indexOf(VATInput.charAt(i))<0){ 
			alert ("Please Fill correct VAT number. e.g 'BE417827795'");
			document.orgForm.orgVat.select(); return false;
			}



//			if(i<2){
//			if (alphabets.indexOf(VATInput.charAt(i))<0){ 
//			alert ("Please Fill correct VAT number. e.g 'BE417827795'");
//			document.orgForm.orgVat.select(); return false;
//			}
//			}
//			
//			else{
//			if (digits.indexOf(VATInput.charAt(i))<0){ 
//			alert ("Please Fill correct VAT number. e.g 'BE417827795'");
//			document.orgForm.orgVat.select(); return false;
//			}
//			}

	 }
	 }
	 }
//'''''''''''''end check

////////////
document.orgForm.Done.value=0;}
////////////////


////////////////////:::
if(Lng=='Fra'){
/////////////////////
 
 		//////case 1
			if(document.orgForm.orgCou.options[opt].value == 550){document.orgForm.VATCheck.value = "YES";}
			//////case 2			
			if((document.orgForm.orgCou.options[opt].value != 550)&&(lnk==1)){
			if(document.orgForm.orgVat.value == ""){ 
			if(confirm("Veuillez compléter votre nombre de TVA. Cliquetez en Cancel pour remplir nombre de TVA . Ou cliquetez Ok si vous n'avez pas un nombre de TVA ."))
			{document.orgForm.VATCheck.value = "YES";}else{return false;}
		    }
		    else{document.orgForm.VATCheck.value = "NO";}
		    }
		    //////case 3
			if(lnk==0){	document.orgForm.VATCheck.value = "NO";}
	
	
//if (c==1) {
//			alert("Vous êtes soumis au Belge TVA.");
//			document.orgForm.VATCheck.value = "YES";
//			if(document.orgForm.orgVat.value == ""){alert("Veuillez compléter votre nombre de TVA."); document.orgForm.orgVat.select(); return false;}
//	      }
//if (d==1) {
//			document.orgForm.VATCheck.value = "YES";
//			if(document.orgForm.orgVat.value == ""){alert("Veuillez compléter votre nombre de TVA."); document.orgForm.orgVat.select(); return false;}
//	      }
//if (e==1) {
//          if(document.orgForm.orgVat.value == ""){document.orgForm.VATCheck.value = "YES";}
//          else{document.orgForm.VATCheck.value = "NO";}
//		  }	
	
//if (f==1) {	    	document.orgForm.VATCheck.value = "YES";		  }	


////////::::to check the format of vat number according to Renoud
	if (document.orgForm.orgVat.value != ""){
		////////////new check for EU country
	if (lnk==1){
	//////////:

	if(document.orgForm.orgVat.value.length<8) {alert ("Veuillez remplir nombre correct de TVA. e.g 'BE417827795'"); document.orgForm.orgVat.select(); return false;}
	
	var alphabets;
	var digits;
	var VATInput;
	alphabets = "ABCDEFGHIJKLMNOPWRSTUVWXYZ0123456789-";
//	alphabets = "ABCDEFGHIJKLMNOPWRSTUVWXYZ";
//	digits="0123456789-";
	VATInput=document.orgForm.orgVat.value;
	for (i = 0; i < VATInput.length; i++){

			if (alphabets.indexOf(VATInput.charAt(i))<0){ 
			alert ("Veuillez remplir nombre correct de TVA. e.g 'BE417827795'");
			document.orgForm.orgVat.select(); return false;
			}
			

	//		if(i<2){
	//		if (alphabets.indexOf(VATInput.charAt(i))<0){ 
	//		alert ("Veuillez remplir nombre correct de TVA. e.g 'BE417827795'");
	//		document.orgForm.orgVat.select(); return false;
	//		}
	//		}
	//			
	//			else{
	//			if (digits.indexOf(VATInput.charAt(i))<0){ 
	//			alert ("Veuillez remplir nombre correct de TVA. e.g 'BE417827795'");
	//			document.orgForm.orgVat.select(); return false;
	//			}
	//			}


	 }
	 }
	 }
//'''''''''''''end check
////////////
document.orgForm.Done.value=0;}
////////////////

////////////////////:::
if(Lng=='Spa'){
/////////////////////

		//////case 1
			if(document.orgForm.orgCou.options[opt].value == 550){document.orgForm.VATCheck.value = "YES";}
			//////case 2			
			if((document.orgForm.orgCou.options[opt].value != 550)&&(lnk==1)){
			if(document.orgForm.orgVat.value == ""){ 
			if(confirm("Please fill the VAT number. Haga clic Cancel para llenar número del IVA . O haga clic OK si usted no tiene un número del IVA."))
			{document.orgForm.VATCheck.value = "YES";}else{return false;}
		    }
		    else{document.orgForm.VATCheck.value = "NO";}
		    }
		    //////case 3
			if(lnk==0){	document.orgForm.VATCheck.value = "NO";}
	
	
//if (c==1) {
//			alert("You are subjected to Belgian VAT.");
//			document.orgForm.VATCheck.value = "YES";
//			if(document.orgForm.orgVat.value == ""){alert("Complete por favor su número del IVA."); document.orgForm.orgVat.select(); return false;}
//	      }
//if (d==1) {
//			document.orgForm.VATCheck.value = "YES";
//			if(document.orgForm.orgVat.value == ""){alert("Complete por favor su número del IVA."); document.orgForm.orgVat.select(); return false;}
//	      }
//if (e==1) {
 //         if(document.orgForm.orgVat.value == ""){document.orgForm.VATCheck.value = "YES";}
  //        else{document.orgForm.VATCheck.value = "NO";}
   //        }	
	
//if (f==1) {	document.orgForm.VATCheck.value = "YES";	 }	


////////::::to check the format of vat number according to Renoud
	if (document.orgForm.orgVat.value != ""){
	
		////////////new check for EU country
	if (lnk==1){
	//////////:

	//if(document.orgForm.orgVat.value.length<8) {alert ("Llene por favor el número correcto del IVA. e.g 'BE417827795'"); document.orgForm.orgVat.select(); return false;}
	
	var alphabets;
	var digits;
	var VATInput;
	alphabets = "ABCDEFGHIJKLMNOPWRSTUVWXYZ0123456789-";
//	alphabets = "ABCDEFGHIJKLMNOPWRSTUVWXYZ";
//	digits="0123456789-";
	VATInput=document.orgForm.orgVat.value;
	for (i = 0; i < VATInput.length; i++){
			
			if (alphabets.indexOf(VATInput.charAt(i))<0){ 
			alert ("Llene por favor el número correcto del IVA. e.g 'BE417827795'");
			document.orgForm.orgVat.select(); return false;
			}

//			if(i<2){
//			if (alphabets.indexOf(VATInput.charAt(i))<0){ 
//			alert ("Llene por favor el número correcto del IVA. e.g 'BE417827795'");
//			document.orgForm.orgVat.select(); return false;
//			}
//			}
//			
//			else{
//			if (digits.indexOf(VATInput.charAt(i))<0){ 
//			alert ("Llene por favor el número correcto del IVA. e.g 'BE417827795'");
//			document.orgForm.orgVat.select(); return false;
//			}
//			}

	 }
	 }
	 }
//'''''''''''''end check
////////////
document.orgForm.Done.value=0;}
////////////////


////////////
}
////////////////
function Assistant(ast)
{
  if (ast==1)
  {wmsg = window.open('vat.htm','PWWnd','width=290,height=170,left=405,top=50,scrollbars=no,status=no,scrollbars=yes');  }
  if (ast==2)
  {wmsg = window.open('how.asp','PWWnd','width=400,height=200,left=405,top=50,menubar=yes,scrollbars=no,status=no,scrollbars=yes');   } 
  wmsg.focus();
}

///////////////////////////:
function CheckPayment()
{
  for (var i=0;i<3;i++){
  if (document.orgForm.payment[i].checked == true) {
    document.orgForm.payment.value=i+2;
    document.orgForm.payment_hid.value=i+2;
   }
  }
}

/////////////////////////

///////////:
function IAgree() 
{
	window.open('../terms.asp','ANWnd','width=700,height=400,left=105,top=50,menubar=yes,scrollbars=yes,status=yes,resize=yes');
}
//////////

function ajax_load(path, containerId, evalCode) {
	//AddBusy(containerId, "Loading", "");
	$.ajax({
		cache: false,
		url: path,
		success: function (result) {
			//RemoveBusy(containerId);
			$('#' + containerId).html(result);
			if (evalCode != "")
				eval(evalCode);
		}
	});
	return false;
}


function ajax_post(formId, path, containerId, evalOnSuccess) {
	//	AddBusy(containerId, "<span class=\"loading\">Loading...</span>", "");

	$.post(path, $("#" + formId).serialize())
		.done(function (result) {
				if (result.indexOf('ERROR:') < 0) {
					$('#' + containerId).html(result);
					if (evalOnSuccess != "")
						eval(evalOnSuccess);
				}
				else 
					$('#' + containerId).html(result.replace('ERROR: ', '<span class=\"error\">') + '</span><br/>' + $('#' + containerId).html());
			})
		.fail(function (a, errStatus, errMsg) {
				$('#' + containerId).html('<span class=\"error\">Error while sending your request.</span><br/>');
			})
		.always(function () {  });

//	$.ajax({
//		cache: false,
//		url: path,
//		type: "POST",
//		enctype: "multipart/form-data",
//		data: $("#" + formId).serialize(),
//		success: function (result) {
//			//	RemoveBusy(containerId);
//			alert(result);
//			if (result.indexOf('ERROR:') < 0) {
//				$('#' + containerId).html(result);
//				if (evalCode != "")
//					eval(evalOnSuccess);
//			}
//			else {
//				$('#' + containerId).html(result.replace('ERROR: ', '<span class=\"error\">') + '</span><br/>' + $('#' + containerId).html());
//			}
	//		}
//		error: function (jqXHR, textStatus, err) {
//			alert('text status ' + textStatus + ', err ' + err)
//		}
//	});

	return false;
}

function AddBusy(containerId, busyTxt, addCssClass) {
	if (containerId != '') {
		$('#' + containerId).html("<div class=\"loadingTxt " + addCssClass + "\">" + (busyTxt != '' ? busyTxt : "Please wait") + "</div>");
	}
}

function RemoveBusy(containerId) {
	$('#' + containerId + ' .loadingTxt').remove();
}
