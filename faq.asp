<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/simpleModal.min.js" type="text/javascript"></script>
<script type="text/javascript">
  $(function () {
    $('#sizeChart').click(function () {
      $.modal('<iframe src="/sizing.asp" height="450" width="550" style="border:0">', {

      });
    });
  });
</script>
<title>FAQ<%=cKeywords_Title%></title>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
<!--#include virtual="/incBanners.asp" -->
    <div class="Content_Area Module"> <!-- START MAIN AREA CONTENT -->
        <div class="Main_Body_LeftColumn_Wide">
            <H1>FAQ</H1>
            
            <div class="faq_redbox">Have different question? Use the <a href="/contact-chestees/">contact form</a> to ask and we'll be sure to respond.</div>
            
            <div class="faq_Q">Q: How do I know what size to buy?</div>
            <div class="faq_A"><b>A:</b> Refer to the <a href="#" id="sizeChart">sizing page</a> for measurement.</div>

            <div class="faq_Q">Q: What kind of shirts do you print on?</div>
            <div class="faq_A"><b>A:</b> We print on <a href="http://store.americanapparel.net/product/index.jsp?productId=2001" target="_blank">American Apparel 2001 Unisex Fine Jersey Short Sleeve shirts</a></div>
          
            <div class="faq_Q">Q: How do you print your shirts?</div>
            <div class="faq_A"><b>A:</b> We hand screen print our shirts in-house using plastisol inks.</div>
            
            <div class="faq_Q">Q: Do you ship outside of the United States?</div>
            <div class="faq_A"><b>A:</b> Not through the website. Use the "<a href="/contact-chestees/">Contact Us</a>" form and we can work it out.</div>
            
            <div class="faq_Q">Q: What forms of payment do you accept?</div>
            <div class="faq_A"><b>A:</b> Visa, Mastercard, Discover and American Express.</div>
            
            <div class="faq_Q">Q: What if I need to make a return?</div>
            <div class="faq_A"><b>A:</b> Do not automatically send it back. Use the "<a href="/contact-chestees/">Contact Us</a>" form and explain the situation.</div>
            
        </div>
        <div class="Main_Body_Random">
            <!--#include virtual="/incRandom.asp" -->
        </div>
        <div class="clear"></div>
    </div> <!-- END MAIN AREA CONTENT -->
    <div class="clear"></div>
</div> <!-- END MAIN AREA -->
<!--#include virtual="/incFooter.asp" -->