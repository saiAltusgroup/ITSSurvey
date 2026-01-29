<!--#INCLUDE virtual="/lib/asp/adovbs.asp"-->
<!--#INCLUDE virtual="/sys_include/db.asp"-->
<!--#INCLUDE virtual="/include/input.asp"-->
<!--#INCLUDE virtual="/include/user.asp"-->

<%

Set Input=new cInput
uid=Input.uid
pwd=Input.pwd
url=Input.url
Set Input=nothing

Set u = new cUser
Call u.Login(uid, pwd)
Set u = Nothing

If (Len(session("uid")) = 0) Then
	append = "?"
	If (InStr(1, url, "?", vbTextCompare) > 0) Then append = "&"

	url = url & append & "failed=1"
End If

Response.Redirect url
%>
