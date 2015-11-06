$(document).ready(function() {
	intProductSizeID = 3;
	$(".Sizes").click(function() {
		intProductSizeID = $(this).attr("id");
		$('.Sizes').removeClass('ON');
		$(this).addClass('ON');
	});
	$("#AddToCart").click(function() {
		$.ajax({
			url: "/submit_Product.asp",
			type: 'post',
			data: $('form').serialize() + '&ProductSizeID=' + intProductSizeID,
			success: function(response) {
				if (response == 'Success') {
					window.location = '/cart/';
				} else {
					$('#colorError').html("<div class='ErrorBar round b_margin_10 txt_center'>Please select a color</div>");
				}
			}}); //END ajax					   
	});
	
	$('#Product_SM_1').click(function() {
		$('#Product_Big').attr("src", $(this).attr("src"));									 
	});
	$('#Product_SM_2').click(function() {
		$('#Product_Big').attr("src", $(this).attr("src"));									 
	});
	$('#Product_SM_3').click(function() {
		$('#Product_Big').attr("src", $(this).attr("src"));									 
	});
	$('#Product_SM_4').click(function() {
		$('#Product_Big').attr("src", $(this).attr("src"));									 
	});
}); //END Document