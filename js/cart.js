$(document).ready(function() {
	
	var dataString = "";
	
	PopulateCart();
	
	function PopulateCart() {
		$("#myCart").html("<div class='loader'><img src='/images/clock-loader.gif'></div>");
		$.ajax({
			url: "/submit_CartUpdate.asp",
			success: function(response) {
				if (response == 'Error') {
					$('#Cart_Updated').css("display", "block");
					$("#Cart_Updated").html("Something went wrong");
				} else {
					$('#myCart').delay(1000).fadeOut('slow', function() {
						$("#myCart").html(response);
						$('#myCart').fadeIn('fast');
						$('input').bind('change', function() {
						 	dataString = dataString + '&' + $(this).attr('name') + '=' + $(this).val();
						});
						BindDelete();
					});
				}
			}}); //END ajax
	}
	
	
	$("#KeepShopping").click(function() {
		window.location = '/';
	});
	
	$("#myCheckout").click(function() {
		window.location = 'https://www.chestees.com/checkout/';
	});
	
	$("#myUpdate").click(function() {
		$("#myCart").html("<div class='loader'><img src='/images/clock-loader.gif'></div>");
		$.ajax({
			url: "/submit_CartUpdate.asp",
			type: 'post',
			data: 'Update=1' + dataString,
			success: function(response) {
				if (response == 'Error') {
					$('#Cart_Updated').css("display", "block");
					$("#Cart_Updated").html("Something went wrong");
				} else {
					dataString = "";
					$("#Cart_Updated").html("Cart Updated");
					$('#Cart_Updated').delay(1000).slideDown('slow');
					$('#myCart').delay(1000).fadeOut('slow', function() {
						$("#myCart").html(response);
						$('#myCart').fadeIn('fast');
						$('input').bind('change', function() {
						 	dataString = dataString + '&' + $(this).attr('name') + '=' + $(this).val();
						});
						BindDelete();
					});
				}
			}}); //END ajax						
	});
	
//	$(".delete").live('click', (function() {
//		//alert("ID: " + $(this).attr('id'));
//		$.ajax({
//			url: "/submit_CartUpdate.asp",
//			type: 'post',
//			data: 'Delete=1&CartID=' + $(this).attr('id'),
//			success: function(response) {
//				if (response == 'Updated') {
//					$('#Cart_Updated').css("display", "block");
//					$("#Cart_Updated").html("Cart Updated");
//					BindDelete()
//				}  else if (response == 'Error') {
//					$('#Cart_Updated').css("display", "block");
//					$("#Cart_Updated").html("Something went wrong");
//				}
//			}}); //END ajax				 
//	}));
	
	function BindDelete() {
		$('.delete').live('click', function() {
			$.ajax({
				url: "/submit_CartUpdate.asp",
				type: 'post',
				data: 'Delete=1&CartID=' + $(this).attr('id'),
				success: function(response) {
					if (response == 'Error') {
					$('#Cart_Updated').css("display", "block");
					$("#Cart_Updated").html("Something went wrong");
				} else {
					$("#Cart_Updated").html("Cart item deleted");
					$('#Cart_Updated').delay(1000).slideDown('slow');
					$('#myCart').delay(1000).fadeOut('slow', function() {
						$('#myCart').fadeIn('fast');
						$("#myCart").html(response);
					});
				}
			}}); //END ajax									
		});
	}
}); //END Document