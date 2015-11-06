<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Contact Chestees<%=cKeywords_Title%></title>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/contact.js" type="text/javascript"></script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
<!--#include virtual="/incBanners.asp" -->
    <div class="Content_Area Module"> <!-- START MAIN AREA CONTENT -->
        <div class="Main_Body_LeftColumn_Wide">
            <H1>Contact Chestees</H1>
            
            <div id="Response" style="display:none;"></div>
            <div id="myForm">
                <div style="margin-bottom:10px;">Your Email Address <span style="font-size:10px;">(so we can respond)</span><br>
                    <input id="Email" name="Email" type="text" class="rounded-glow w270">
                </div>          
                <div style="margin-bottom:10px; padding-bottom:10px; border-bottom:1px dashed #b93636;">Say what you have to say. Ask what you need to ask.<br>
                    <textarea id="Comment" name="Comment" class="formTextarea" style="width:380px; height:200px;"></textarea>
                </div>
                <div style="margin-bottom:10px;"><input name="Submit" id="mySubmit" class="medium red button" type="submit" value="Send my message"></div>              
        	</div>
        </div>
        <div class="Main_Body_Random">
            <!--#include virtual="/incRandom.asp" -->
        </div>
        <div class="clear"></div>
    </div> <!-- END MAIN AREA CONTENT -->
    <div class="clear"></div>
</div> <!-- END MAIN AREA -->
<!--#include virtual="/incFooter.asp" -->