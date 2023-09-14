function checkNumeric(obj_field, error_message, is_focused) {
	if (!(obj_field)) { return true; }
	if (isNaN(obj_field.value)) {
		if (error_message.length>0) {
			alert(error_message);
		}
		if((is_focused || is_focused==1) && obj_field.focus) {
			obj_field.focus()
		}
		return false;
	}
}

function checkTextFieldValue(obj_field, field_value, error_message, is_focused) {
	if (!(obj_field)) { return true; }
	if (obj_field.value==field_value) {
		if (error_message.length>0) {
			alert(error_message);
		}
		if((is_focused || is_focused==1) && obj_field.focus) {
			obj_field.focus()
		}
		return false;
	}
	return true;
}

function checkTextFieldLength(obj_field, field_length, error_message, is_focused) {
	if (!(obj_field)) { return true; }
	if (obj_field.value.length>=field_length) {
		if (error_message.length>0) {
			alert(error_message);
		}
		if((is_focused || is_focused==1) && obj_field.focus) {
			obj_field.focus()
		}
		return false;
	}
	return true;
}

function checkSelectFieldIndex(obj_field, field_index, error_message, is_focused) {
	if (!(obj_field)) { return true; }
	if (obj_field.selectedIndex==field_index) {
		if (error_message.length>0) {
			alert(error_message);
		}
		if((is_focused || is_focused==1) && obj_field.focus) {
			obj_field.focus()
		}
		return false;
	}
	return true;
}

function checkDateComposition(y, m, d, error_message) {
	if (isNaN(y) || isNaN(m) || isNaN(d)) { return true; }

	var test_date=new Date();
	test_date.setFullYear(y, (m-1), d);
	if (test_date.getMonth()==(m-1) && test_date.getDate()==d) {
		return true; }
	else {
		if (error_message.length>0) {
			alert(error_message);
		}
		return false;}
}

function validateEmail1(email) {
	var emailPattern = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,5}$/; 
	return emailPattern.test(email); 
} 

function validateEmail(email) {
	var emails=email.split(";")
	var result = true;
	l=emails.length;
	for (i=0;i<l;i++) {
		result = (result && validateEmail1(emails[i].replace(" ", "")));
	}
	return result;
} 
