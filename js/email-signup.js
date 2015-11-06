$(function () {
  $("#btnSignUp").click(function () {
    var email = $("input#txtEmailSignUp").val();
    $(".EmailSignUp.Module #EmailSignUp_Input").html("<div class='loader'><img src='/images/ajax-loader-white.gif'></div>");
    $("#EmailSignUp_Text").empty();
    $.ajax({
      url: "/submit_Email.asp",
      type: 'post',
      data: "Email=" + email,
      success: function (response) {
        $('#EmailSignUp_Text').fadeOut('fast');
        if (response == 'Success') {
          $('#EmailSignUp_Title').fadeOut('fast');
          $('#EmailSignUp_Input').fadeOut('fast', function () {
            $("#EmailSignUp_Input").empty();
            $("#EmailSignUp_Input").append("<div style='font-size:14px;'>Thanks for signing up!<br /><br />Have you joined us on <a href='http://www.facebook.com/chestees' target='_blank'>Facebook</a> yet?<br /><br />~Love Chestees</div>").fadeIn('fast');
          })
        } else {
          $('#EmailSignUp_Text').fadeOut('fast', function () {
            $("#EmailSignUp_Input").empty();
            $("#EmailSignUp_Text").append("<div style='font-size:14px;'>Uh Oh...Something wasn't right about that email.<br /><br />Try that again.</div>").fadeIn('fast');
          })
        }
      } 
    }); //END ajax
  }); //END Button Click		
}); //END Document