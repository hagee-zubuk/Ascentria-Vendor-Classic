function Left(str, n) {
	if (n <= 0)
	    return "";
	else if (n > String(str).length)
	    return str;
	else
	    return String(str).substring(0,n);
}
function RTrim(str) {
	var whitespace = new String(" \t\n\r");
	var s = new String(str);
	if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
		var i = s.length - 1;       
		while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1) i--;

		s = s.substring(0, i+1);
	}
	return s;
}
function LTrim(str) {
	var whitespace = new String(" \t\n\r");
	var s = new String(str);
	if (whitespace.indexOf(s.charAt(0)) != -1) {
		var j=0, i = s.length;
		while (j < i && whitespace.indexOf(s.charAt(j)) != -1) j++;
		
		s = s.substring(j, i);
	}
	return s;
}
function Trim(str) {
	return RTrim(LTrim(str));
}
function maskMe(str,textbox,loc,delim) {
	var locs = loc.split(',');
	for (var i = 0; i <= locs.length; i++) {
		for (var k = 0; k <= str.length; k++) {
			if ( (k == locs[i]) && (str.substring(k, k+1) != delim) ) {
				str = str.substring(0,k) + delim + str.substring(k,str.length);
			}
		}
	}
	textbox.value = str
}
function bawal(tmpform) {
	var iChars = ",|\"\'";
	var tmp = "";
	for (var i = 0; i < tmpform.value.length; i++) {
		if (iChars.indexOf(tmpform.value.charAt(i)) != -1) {
			alert ("This character is not allowed.");
			tmpform.value = tmp;
			return;
		} else {
			tmp = tmp + tmpform.value.charAt(i);
		}
	}
}
function bawal2(tmpform) {
	var iChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz0123456789-,.\'"; //",|\"\'";
	var tmp = "";
	for (var i = 0; i < tmpform.value.length; i++) {
		if (iChars.indexOf(tmpform.value.charAt(i)) != -1) {
			tmp = tmp + tmpform.value.charAt(i);
		} else {
			alert ("This character is not allowed.");
			tmpform.value = tmp;
			return;
		}
	}
}
function bawalletters(tmpform) {
	var iChars = "0123456789";
	var tmp = "";
	for (var i = 0; i < tmpform.value.length; i++) {
		if (iChars.indexOf(tmpform.value.charAt(i)) != -1) {
			tmp = tmp + tmpform.value.charAt(i);
		} else {
			alert ("This character is not allowed.");
			tmpform.value = tmp;
			return;
		}
	}
}
