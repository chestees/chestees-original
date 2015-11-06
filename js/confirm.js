$(document).ready(function() {
	$('#mySubmit').click(function() {
		$('#MessageBar').empty();
		
		var PurchaseAmount = $("input#PurchaseAmount").val();
		var ShippingCost = $("input#ShippingCost").val();
		var CouponCode = $("input#CouponCode").val();
		var TotalAmount = $("input#TotalAmount").val();
		$.ajax({
			url: "/submit_Confirm.asp",
			type: 'post',
			data: 'PurchaseAmount='+PurchaseAmount+'&ShippingCost='+ShippingCost+'&CouponCode='+CouponCode+'&TotalAmount='+TotalAmount,
			dataType: "html",
			async: false,
			error: function() {
				//alert("ERROR");	
			},
			success: function(Response) {
				if (Response == 'Success') {
					window.location = 'https://www.chestees.com/thank-you/';
				} else {
					$('#MessageBar').append("<div class='ErrorBar'></div>");
					$(window).scrollTop(0);
					strError = "<div style='width:850px; margin:0 auto;'>"
					strError = strError + "<div class='ErrorBar_L'><img src='/images/icon-error-lg.png' width='39' height='35'></div>"
					strError = strError + "<div class='ErrorBar_R'>Oh No! "+Response+"</div>";
					strError = strError + "<div class='clear'></div></div>";
					$(".ErrorBar").html(strError);
					$('#MessageBar').delay(500).slideDown('slow');
				}
			}}); //END ajax
	});
}); //END Document