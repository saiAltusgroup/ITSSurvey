<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/sys_include/globals.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->

<%
Set Input=new cInput
site=Input.site
page=Input.page
Set Input=nothing


form_action= ""
form_name = "login"

Set u = new cUser
userid=u.UserId
password=u.GetPassword()
isloggedin=u.IsLoggedIn
Set u = Nothing

Select Case site
	Case "main"
		targetpage = "index.php"
		if (page <> "") then targetpage = page
		form_action = g_SITE_MAIN & targetpage
	Case "about"
		targetpage = "index.php"
		if (page <> "") then targetpage = page
		form_action = g_SITE_ABOUT & targetpage
	Case "office"
		form_action = g_SITE_OFFICE & "login.asp"
		form_name = "gologin"
	Case "industrial"
		form_action = g_SITE_INDUSTRIAL & "index.php"
		form_name = "form99"
	Case "retail"
		form_action = g_SITE_RETAIL & "index.php"
		form_name = "form99"
	Case "investment"
		form_action = g_SITE_INVESTMENT & "login.asp"
		if (NOT isloggedin) then form_action = g_SITE_INVESTMENT & "eng/chart/default.asp"
		If (Len(page) = 0) Then page = "default.asp"
	Case "myreports"
		form_action = g_SITE_MYREPORTS & "main/loginAction.asp"
		form_name = "form99"
	Case Else
		html_base = Request.ServerVariables("DOCUMENT_ROOT") & "/"
		Response.Redirect(html_base & "default.asp")
End Select

If (isloggedin) Then
ElseIf ((site = "main") AND (form_action = "http://www.altusinsite.com//index.php?page=contact") or form_action = "http://ai.gmont.ca//index.php?page=contact") Then
	Response.Redirect(form_action )
ElseIf ((site = "main") AND (form_action = "http://www.altusinsite.com//fr/index.php?page=contact" or form_action = "http://ai.gmont.ca//fr/index.php?page=contact")) Then
   	Response.Redirect(form_action )
ElseIf ((site <> "office") AND (site <> "investment")) Then
	Response.Redirect(form_action & "?url=" & page & "&page=" & page)
ElseIf (Site = "investment") Then
	Response.Redirect(form_action & "?url=" & page & "&page=" & page)
Else
	form_action = g_SITE_OFFICE & "default.asp"
End If

If isloggedin Then
%>

	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
	<html><head><title>Cross Redirect</title></head><body>
	<form name="<%=form_name%>" id="<%=form_name%>" action="<%=form_action%>" method="post">
	<input type="hidden" name="form99_action" value="login">
	<input type="hidden" name="source" value="form99">
	<input type="hidden" name="url" value="<%=page%>">
	<input type="hidden" name="uid" value="<%=userid%>">
	<input type="hidden" name="pwd" value="<%=password%>">
	<input type="hidden" name="userid" value="<%=userid%>">
	<input type="hidden" name="passwd" value="<%=password%>">
	<input type="hidden" name="form99_userid" value="<%=userid%>">
	<input type="hidden" name="form99_password" value="<%=password%>">
	<input type="hidden" name="referingsite" value="investment">
	</form>
	<script type="text/javascript">
	<!-- to hide script contents from old browsers
	document.<%=form_name%>.submit();
	// end hiding contents from old browsers -->
	</script>
	</body></html>

<%
Else
	Response.Redirect(form_action)
End If
%>
