(function($) {
	"use strict";
	var HT = {}; // Khai báo là 1 đối tượng

	/* MAIN VARIABLE */

	var $window = $(window),
	    $document = $(document),
		$carousel = $(".owl-slide");
	    

	    // FUNCTION DECLARGE
	    $.fn.elExists = function() {
	        return this.length > 0;
	    };
		HT.carousel = () => {
			$carousel.each(function(){
				let _this = $(this);
				let option = _this.find('.owl-carousel').attr('data-owl');
				let owlInit = atob(option);
				owlInit = JSON.parse(owlInit);
				_this.find('.owl-carousel').owlCarousel(owlInit);
			});
			
		} 
		

	    
	    // Document ready functions
	    $document.on('ready', function() {
	    	HT.carousel();
	    });

	})(jQuery);
	$(function () {
		$(document).scroll(function () {
		  var $nav = $(".navbar-fixed-top");
		  $nav.toggleClass('scrolled', $(this).scrollTop() > $nav.height());
		});
	  });
	  $(document).ready(function() {
		var wd_width = $(window).width();
		if(wd_width > 1220) {
			new WOW().init();
		}
	});