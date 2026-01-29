$(document).ready(function() {

	$('#Definitions dd').addClass('collapsed'); // hide definitions
	$('#Definitions dt').addClass('unselected'); // Set all dt's to the collapse displayed default

	// add onclick events to the dt's
	$('#Definitions dt').click(function () {

		currentClass = $(this).attr('class')
		if(currentClass == 'selected') { // check for current class
			$(this).removeClass('selected'); //remove selected formating
			$(this).addClass('unselected'); //change back to the default
			$(this).next("dd").removeClass('selected'); //fix for older browsers
		}
		else {
			$(this).removeClass('unselected'); // remove the default
			$(this).addClass('selected'); //add selected formating
			$(this).next("dd").addClass('selected'); //fix for older browsers
		}

		$(this).next("dd").slideToggle(500); //toggle the display

	});

});