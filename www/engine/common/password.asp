<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/language.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/page.asp"-->
<!--#INCLUDE virtual="/include/converter.asp"-->

<%

Set Input=new cInput
step=Input.step
email=Input.email
Set Input=nothing

Set p=new cPage
p.Title=Lang("title_password_retrieval")
p.Subhead=p.Title
p.Protect=false


result = true

If (step = 1) Then

	Set u = new cUser
	result = u.SendPassword(email)
	Set u = Nothing

	If (result) Then
		Call p.Display("ContentThankYou", true)
	Else
		Call p.Display("ContentPasswordRetrieval", true)
	End If

Else
	Call p.Display("ContentPasswordRetrieval", true)
End If

Set p = nothing


Function ContentPasswordRetrieval
%>
	<br>
	<form style="width: 350px;" class="objcenter searchform" name="form98" id="form98" method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>">
		<fieldset>
			<legend><%=Lang("password_retrieval")%></legend>
		<%If result = false Then%>
			<p class="objcenter" style="width: 244px; color: #f00;"><%=Lang("error_email_notfound")%></p>
		<%End If%>
			<p><%=Lang("intro_email")%></p>
			<label for="email"><%=Lang("email_label")%></label>
			<input name="email" id="email" size="40" maxlength="60" type="text">
			<input type="hidden" name="step" value="1">
			<div class="textcenter">
				<input class="search" type="submit" name="action" id="action" value="<%=Lang("check")%>">
			</div>
		</fieldset>
	</form>
	<br><br><br>
<%
End Function

Function ContentThankYou
%>
	<br>
	<form style="width: 350px;" class="objcenter searchform" name="form98" id="form98" >
		<fieldset>
			<legend><%=Lang("password_retrieval")%></legend>
			<p><%=Lang("info_email_found")%><br><br>
			<a href="/default.asp"><%=Lang("link_login")%></a></p>
		</fieldset>
	</form>
	<br><br><br>
<%
End Function
%>
