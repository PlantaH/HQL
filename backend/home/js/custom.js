

/*=============================================================
    Authour URI: www.binarytheme.com
    License: Commons Attribution 3.0

    http://creativecommons.org/licenses/by/3.0/

    100% To use For Personal And Commercial Use.
    IN EXCHANGE JUST GIVE US CREDITS AND TELL YOUR FRIENDS ABOUT US
   
    ========================================================  */


(function ($) {
    "use strict";
    var mainApp = {

        main_fun: function () {

            //  SCROLL SCRIPT FUNCTION FOR NAVBAR 
         $(function () {
             $('.scrollclass a').bind('click', function (event) { //just pass scrollclass in design and start scrolling
                 var $anchor = $(this);
                 $('html, body').stop().animate({
                     scrollTop: $($anchor.attr('href')).offset().top
                 }, 1000, 'easeInOutQuad');
                 event.preventDefault();
             });
         });
            //ADD REMOVE PADDING CLASS ON SCROLL
            $(window).scroll(function () {
                if ($(".navbar").offset().top >40) {
                    $(".navbar-fixed-top").addClass("navbar-pad-original");
                } else {
                    $(".navbar-fixed-top").removeClass("navbar-pad-original");
                }
            });
       


            /*====================================
               WRITE YOUR SCRIPTS BELOW 
           ======================================*/


        },

        initialization: function () {
            mainApp.main_fun();

        }

    }
    // Initializing ///

    $(document).ready(function () {
        mainApp.main_fun();
    });

}(jQuery));

//Only numbers
//NOTE Inline: call to from code line in HTML node, ej: onkeypress="return isNumberKey(event)"
function isNumberKeyInLine(evt){
	var charCode = (evt.which) ? evt.which : evt.keyCode;
	if (charCode != 46 && charCode > 32 
	&& (charCode < 48 || charCode > 57))
	return false;
	return true;
} 
//

//
function isNumberKey(evt){
	console.log(evt);
	var charCode = (evt.which) ? evt.which : evt.keyCode;
	if ((charCode < 48 || charCode > 57)
	&& (charCode < 96 || charCode > 105)
	&& (charCode >= 32)
	&& (charCode < 48 || charCode > 57)
	&& (charCode != 8))
	return false;
	return true;
} 
//

//Only numbers and only one dot (.)
function isNumberAndDotKey(node, evt){
	var charCode = (evt.which) ? evt.which : evt.keyCode,
		thisNode = node,
		thisNodeVal = thisNode.value,
		dotExist = thisNode.value.indexOf('.') !== -1 ? 'yes' : 'not';
		//console.log(charCode);
	if ((charCode != 46 || (charCode == 46 && dotExist == 'yes')
	&& charCode > 32)
	&& (charCode < 48 || charCode > 57)
	&& (charCode != 8))
	return false;
	return true;
} 
//



