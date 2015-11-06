function isValidDate(dateStr) {
	// Checks for the following valid date formats:
	// MM/DD/YY   MM/DD/YYYY   MM-DD-YY   MM-DD-YYYY
	// Also separates date into month, day, and year variables

	var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{2}|\d{4})$/;

	// To require a 4 digit year entry, use this line instead:
	// var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{4})$/;

	var matchArray = dateStr.match(datePat); // is the format ok?
	if (matchArray == null) {
		//alert("Please Enter a Valid Date.");
		return false;
    }
	month = matchArray[1]; // parse date into variables
	day = matchArray[3];
	year = matchArray[4];
	  
	if (month < 1 || month > 12) { // check month range
		//alert("Please Enter a Valid Date.");
		return false;
	}
	  
	if (day < 1 || day > 31) {
		//alert("Please Enter a Valid Date.");
		return false;
	}

	if ((month==4 || month==6 || month==9 || month==11) && day==31) {
		//alert("Please Enter a Valid Date.");
		return false
	}
	  
	if (month == 2) { // check for february 29th
		var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
		if (day>29 || (day==29 && !isleap)) {
			//alert("Please Enter a Valid Date.");
			return false;
		}
	}
		
	return true;  // date is valid
}
