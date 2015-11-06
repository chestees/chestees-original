$(document).ready(function() {
	$('#mySubmit').click(function() {
		$('#MessageBar').empty();
		$("input:text").each(function (i) {
			if (this.style.background = "#FF6") {
			  this.style.background = "none";
			}
		});
		$("select").each(function (i) {
			if (this.style.background = "#FF6") {
			  this.style.background = "none";
			}
		});
		$.ajax({
			url: "/submit_Checkout.asp",
			type: 'post',
			data: $('form').serialize(),
			dataType: "html",
			async: false,
			error: function() {
				//alert("ERROR");	
			},
			success: function(msg) {
				if (msg == 'Success') {
					window.location = 'https://www.chestees.com/confirm/';
				} else {
					$('#MessageBar').append("<div class='ErrorBar'></div>");
					$(window).scrollTop(0);
					strError = "<div style='width:850px; margin:0 auto;'>"
					strError = strError + "<div class='ErrorBar_L'><img src='/images/icon-error-lg.png' width='39' height='35'></div>"
					strError = strError + "<div class='ErrorBar_R'>Required fields were left blank or the data is invalid. Please correct the highlighted fields.</div>";
					strError = strError + "<div class='clear'></div></div>";
					$(".ErrorBar").html(strError);
					$('#MessageBar').delay(1000).slideDown('slow', function() {
						var mySplitResult = msg.split(",");
						jQuery.each(mySplitResult, function() {
							var myDiv = '#' + this;
							$(myDiv).css('background','#FF6');
					   });
					});
				}
			}}); //END ajax
	});
}); //END Document