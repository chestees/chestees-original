<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include virtual="/includes/globalLib.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<title>Create Chestees Account<%=cKeywords_Title%></title>
<script src="/js/jquery-1.4.2.min.js" type="text/javascript"></script>
<script src="/js/jquery-ui-1.8.2.custom.min.js" type="text/javascript"></script>
<script src="/js/createAccount.js" type="text/javascript"></script>
<!--#include virtual="/incHeader.asp" -->
<div class="Main">
<!--#include virtual="/incBanners.asp" -->
    <div class="Content_Area Module"> <!-- START MAIN AREA CONTENT -->
        <div class="Main_Body_LeftColumn_Wide">
            <H1>Be a Chestees Member</H1>
            
            <div id="Response" style="display:none;"></div>
            
            <div id="myForm">
                <div style="margin-bottom:10px;"><b>Account info</b><br>
                    Email addresses are used soley by chestees.com and are not sold or distrubuted to any outside source.</div>
                
                <div style="margin-bottom:10px;">Email Address<br>
                    <input id="Email" type="text" class="rounded-glow w270">
                </div>
                <div style="margin-bottom:10px;">Confirm Email Address<br>
                    <input id="Email_Confirm" type="text" class="rounded-glow w270">
                </div>
                
                <div style="margin-bottom:10px;">Password<br>
                    <input id="Password" type="password" class="rounded-glow w150">
                </div>
                
                <div style="margin-bottom:10px;">Confirm Password<br>
                    <input id="Password_Confirm" type="password" class="rounded-glow w150">
                </div>
                
                <div style="margin-bottom:10px;">How did you discover Chestees.com?<br>
                    <select id="How" size="1" class="rounded-glow">
                    <option value="">_________________</option>
                    <option value="Search Engine">Search Engine</option>
                    <option value="Friend">Friend</option>
                    <option value="Digg.com">Digg.com</option>
                    <option value="Facebook">Facebook</option>
                    <option value="Dumb Luck">Dumb Luck</option>
                    </select></div>
                <div style="margin-bottom:10px;">If other...<br><input id="How_Other" type="text" class="rounded-glow w270"></div>
                <div style="margin-bottom:10px; padding-bottom:10px; border-bottom:1px dashed #b93636;"><input type="checkbox" id="OptIn">&nbsp;&nbsp;Sign me up to recieve e-mail updates</div>
                <div><input id="mySubmit" class="medium red button" type="submit" value="Join Chestees"></div>
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