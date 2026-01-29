<!--#include virtual="/sys_include/contactform.config.asp"-->
<%
if(Request("lan") <> "") then lan = Request("lan") else lan = "english"
%>
<!--#include file="dict_english.asp"-->
<%
post = false
if(Request("Type") <> "") then strType = Request("Type") else strType = ""
if(Request("Name") <> "") then name = Request("Name") else name = ""
if(Request("CompanyName") <> "") then company_name = Request("CompanyName") else company_name = ""
if(Request("Phone") <> "") then phone = Request("Phone") else phone = ""
if(Request("Email") <> "") then email = Request("Email") else email = ""
if(Request("Comments") <> "") then comments = Request("Comments") else comments = ""
if(Request("Submit") <> "" and (Request("Submit") = LangDict.Item("Submit"))) then post = true

if(post) then
	valid = true
	message = "<ul id=""ErrorMessage"">"

	body = ""

	if(len(strType) <= 0) then
		valid = false
		message = message & "<li>" & LangDict.Item("Legend2") & " " & LangDict.Item("Notprovided") & "</li>"
	end if

	if(len(name) <= 0) then
		valid = false
		message = message & "<li>" & LangDict.Item("Name") & " " & LangDict.Item("Notprovided") & "</li>"
	end if

	if(len(company_name) <= 0) then
		valid = false
		message = message & "<li>" & LangDict.Item("CompanyName") & " " & LangDict.Item("Notprovided") & "</li>"
	end if

	if(len(phone) <= 0) then
		valid = false
		message = message & "<li>" & LangDict.Item("YourPhone") & " " & LangDict.Item("Notprovided") & "</li>"
	else
		if (NOT val_must_be_valid_phone(phone)) then
			valid = false
			message = message & "<li>" & LangDict.Item("YourPhone") & " " & LangDict.Item("Invalid") & "</li>"
		end if
	end if

	if(len(email) <= 0) then
		valid = false
		message = message & "<li>" & LangDict.Item("YourEmail") & " " & LangDict.Item("Notprovided") & "</li>"
	else
		if (NOT val_must_be_email_address(email)) then
			valid = false
			message = message & "<li>" & LangDict.Item("YourEmail") & " " & LangDict.Item("Invalid") & "</li>"
		end if
	end if

	if(instr(strType,"General") > 0 or instr(strType,"Data Update") > 0 or instr(strType,"Support") > 0) then
		if(len(comments) = 0) then
			valid = false
			message = message & "<li>" & LangDict.Item("YourMessage") & " " & LangDict.Item("Notprovided") & "</li>"
		end if
	end if

	body = body & LangDict.Item("About") & " " & strType & vbNewLine
	body = body & LangDict.Item("Name") & " " & Server.HTMLEncode(name) & vbNewLine
	body = body & LangDict.Item("CompanyName") & " " & Server.HTMLEncode(company_name)  & vbNewLine
	body = body & LangDict.Item("YourPhone") & " " & Server.HTMLEncode(phone) & vbNewLine
	body = body & LangDict.Item("YourEmail") & " " & Server.HTMLEncode(email) & vbNewLine
	body = body & LangDict.Item("YourMessage") & " " & Server.HTMLEncode(comments) & vbNewLine
	message = message & "</ul>"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
  <head>
    <title><%=LangDict.Item("PageTitle")%></title>
    <style type="text/css" media="screen" >
      <!--
        @import url('contactform.css');
      -->
    </style>
		<!--[if IE]>
				<link rel="stylesheet" type="text/css" href="contactform_ie.css" />
		<![endif]-->
  </head>
  <body>
		 <h1><img src="altus.jpg" alt="Altus insite" /></h1>
<%
if(post) then
	if(valid) then
	
		if(strType ="Support") then
			strTo = toSupport
		elseif(strType = "Subscription" or strType = "Training/Orientation") then
			strTo = toSubscribe
		else
			strTo = toGeneral
		end if

		Set Mailer = CreateObject("CDONTS.Newmail")
		if isobject(Mailer)=false then
			Response.Write("asp blows")
		end if
		Mailer.From = from
		Mailer.Subject = "Contact Form: " & product
		'Mailer.To = "its@altusinsite.com"
		Mailer.To = strTo
		Mailer.Body = body
		Mailer.Value("Reply-To") = email

		Mailer.Send

		Response.Write(LangDict.Item("Success"))

		set Mailer = nothing
		Response.Write("<br /><br /><br /><br /><a href=""#"" onclick=""javascript:window.close();return false;"">" & LangDict.Item("CloseWindow") & "</a>")
	else
		Response.Write(message)
	end if
end if
if((post and Not(valid)) or Not(post)) then
%>
	<form id="Contact" action="default.asp" method="post">
		<div>
		<input type="hidden" id="lan" name="lan" value="<%=lan%>" />
		<fieldset>
			<legend><%=LangDict.Item("Legend1")%></legend>
			<fieldset>
				<legend><%=LangDict.Item("Legend2")%></legend>
				<ol>
					<li class="checkbox"><input type="checkbox" value="Support" name="Type" id="type05" /><label for="type05"><%=LangDict.Item("Type05")%></label></li>
					<li class="checkbox"><input type="checkbox" value="Subscription" name="Type" id="type02" /><label for="type02"><%=LangDict.Item("Type02")%></label></li>
					<li class="checkbox"><input type="checkbox" value="Training/Orientation" name="Type" id="type03" /><label for="type03"><%=LangDict.Item("Type03")%></label></li>
					<li class="checkbox"><input type="checkbox" value="Data Update" name="Type" id="type04" /><label for="type04"><%=LangDict.Item("Type04")%></label></li>
					<li class="checkbox"><input type="checkbox" value="General" name="Type" id="type01" /><label for="type01"><%=LangDict.Item("Type01")%></label></li>
				</ol>
			</fieldset>
			<br style="clear:both" />
			<ol>
				<li class="textbox">
					<label for="Name"><%=LangDict.Item("Name")%></label><br />
					<input id="Name" name="Name" type="text" size="80" value="<%=name%>" />
				</li>
				<li class="textbox">
					<label for="CompanyName"><%=LangDict.Item("CompanyName")%></label><br />
					<input id="CompanyName" name="CompanyName" type="text" size="80" value="<%=company_name%>" />
				</li>
				<li class="textbox">
					<label for="Phone"><%=LangDict.Item("YourPhone")%></label><br />
					<input id="Phone" name="Phone" type="text" size="80" value="<%=phone%>" />
				</li>
				<li class="textbox">
					<label for="Email"><%=LangDict.Item("YourEmail")%></label><br />
					<input id="Email" name="Email" type="text" size="80" value="<%=email%>" />
				</li>
				<li class="textarea">
					<label for="Comments"><%=LangDict.Item("YourMessage")%></label><br />
					<textarea id="Comments" name="Comments" rows="5" cols="60" ><%=comments%></textarea>
				</li>
				<li class="buttons"><p><input name="Submit" type="submit" value="<%=LangDict.Item("Submit")%>" /><input name="reset" type="reset" value="<%=LangDict.Item("Reset")%>" /></p></li>
			</ol>
		</fieldset>
		</div>
	</form>

<%
end if
%>
</body>
</html>
<%
Function val_must_be_valid_phone(Value)
	Dim ValidationSuccessful
	ValidationSuccessful = false

	Dim regEx, retVal
	Set regEx = New RegExp
	Value = Trim(Value)
	Value = Replace(Value, " ", "")
	Value = Replace(Value, "-", "")
	Value = Replace(Value, "+", "")
	Value = Replace(Value, "(", "")
	Value = Replace(Value, ")", "")

	' Create regular expression:
	regEx.Pattern = "^[\d]{10,11}$"

	' Set pattern:
	regEx.IgnoreCase = true

	' Set case sensitivity.
	retVal = regEx.Test(Value)

	' Execute the search test.
	If retVal Then
		ValidationSuccessful = true
	End If

	val_must_be_valid_phone = ValidationSuccessful
End Function

Function val_must_be_email_address(Value)
	Dim ValidationSuccessful
	ValidationSuccessful = false

	Dim regEx, retVal
	Set regEx = New RegExp

	' Create regular expression:
	regEx.Pattern = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,4}$"

	' Set pattern:
	regEx.IgnoreCase = true

	' Set case sensitivity.
	retVal = regEx.Test(Trim(Value))

	' Execute the search test.
	If retVal Then
		ValidationSuccessful = true
	End If

	val_must_be_email_address = ValidationSuccessful
End Function
%>