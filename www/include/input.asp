<%

' DK, Apr 01, 2008 - version 1.0

Class cInput

	' For reports

	Public Property Get step()
		step=get_request_clng("step")
		if step<0 or step>2 then step=0
	End Property

	Public Property Get period()

		'Set u=new cUser
		'superuser=(u.AdminLevel=10)
		'Set u=nothing

		period=0

		' Note: before only superusers could choose custom period, all other users are restricted to g_CURRENT_PERIOD_ID
		'if superuser then
		period=get_request_clng("period")
		if period<=0 or period>g_MAX_PERIOD_ID then period=g_CURRENT_PERIOD_ID


	End Property

	Public Property Get market()
		market=get_request_arr_clng(find_field("market"))
	End Property

	Public Property Get report()
		report=get_request(find_field("report"))
	End Property

	Public Property Get product()
		product=get_request_clng(find_field("product"))
		if product<0 then product=0
	End Property

	Public Property Get property()
		property=get_request_arr_clng(find_field("property"))
	End Property

	Public Property Get histlen()
		histlen=get_request_clng(find_field("histlen"))
		if histlen<>0 then histlen=5
	End Property


	' Security related

	Public Property Get uid()
		uid=get_request("uid")
	End Property

	Public Property Get pwd()
		pwd=get_request("pwd")
	End Property

	Public Property Get url()
		url=get_request("url")
	End Property

	Public Property Get login()
		login=get_request("login")
	End Property

	' Ajax related

	Public Property Get change()
		change=get_request("change")
	End Property

	Public Property Get firstid()
		firstid=get_request("firstid")
	End Property

	Public Property Get secondid()
		secondid=get_request_clng("secondid")
	End Property


	' Common

	Public Property Get action()
		action=get_request("action")
	End Property

	Public Property Get email()
		email=get_request("email")
	End Property

	Public Property Get site()
		site=get_request("site")
		if "site"="" then site="unknown"
	End Property

	Public Property Get page()
		page=get_request("page")
		if "page"="" then page=""
	End Property

	Public Property Get format()
		format=get_request_clng("format")
	End Property




	' string '
	Private Function get_request(item)
		s=Request(item)
		get_request=s
	End Function

	' int '
	Private Function get_request_clng(item)
		s=Request(item)
		get_request_clng=0
		if isNumeric(s) then get_request_clng=CLng(s)
	End Function

	' array of ints '
	Private Function get_request_arr_clng(item)

		s=Request(item)

		' security check

		tmp_arr=split(s, ",")

		s=""
		comma=""

		for i=0 to ubound(tmp_arr)
			if isnumeric(tmp_arr(i)) then
				s = s & comma & CLng(tmp_arr(i))
				comma=","
			end if
		next

		get_request_arr_clng=s

	End Function

	Private Function find_field(field)

		for i=2 to 0 step -1
			find_field=field & "_" & i
			if Len(Request(find_field))>0 Then Exit Function
		next

		find_field=field

	End Function

End Class

%>
