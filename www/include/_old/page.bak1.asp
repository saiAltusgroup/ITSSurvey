<%

' DK, Apr 01, 2008 - version 1.0

Class cPage

	Public Title
	Public Subhead
	Public BodyAttribs
	Public Protect
	Public Include_head
	Public Include_its
	Public Include_about
	Public Include_cascade
	Public Include_pdf

	Public Base_acs
	Public Base_its

	Public UserObj


	' Constructor
	Private Sub Class_Initialize()
		' Set defaults
		Title=Lang("title_default")
		Subhead=""
		BodyAttribs=""
		Protect=true

		Include_head=true
		Include_its=true
		Include_about=false
		Include_cascade=false
		Include_pdf=false

		Base_acs="http://acs.realinsite.com"
		Base_its="http://" & Request.ServerVariables("SERVER_NAME")

		Set UserObj = new cUser
	End Sub


	' Destructor
	Private Sub Class_Terminate()
		Set UserObj = Nothing
	End Sub


	Public Function Display(ContentArea, sidebar_area)
		Display=DisplayEx(ContentArea, sidebar_area, true)
	End Function

	Public Function GetDisplay(ContentArea, sidebar_area)
		GetDisplay=DisplayEx(ContentArea, sidebar_area, false)
	End Function


	Private Function DisplayEx(ContentArea, sidebar_area, nobuffer)

		DisplayEx=""
		nobuffer_content=nobuffer

		if Protect and (not UserObj.IsLoggedIn) then

			sidebar_area=true

			ContentArea = "GetLoginBox"
			nobuffer_content=false

			BodyAttribs=""

			Include_about=false
			Include_cascade=false
			Include_pdf=false

		end if

		Set ContentFunction = GetRef(ContentArea)


		urlvalues=Request.ServerVariables("QUERY_STRING")

		separator=""
		if Len(urlvalues)>0 then separator="&"

		for each thing in Request.Form

			If Request.Form(thing)<>"" Then
				urlvalues = urlvalues & separator & thing & "=" & Request.Form(thing)
				separator = "&"
			End If
		next

		url=Request.ServerVariables("SCRIPT_NAME") & Server.HTMLEncode("?format=1&" & urlvalues)


		s = ""

		s = s & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbNewLine
		s = s & "<html lang=""" & Lang("lang") & """>" & vbNewLine
		s = s & "<head>" & vbNewLine

		s = s & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbNewLine
		s = s & "<title>" & Title & "</title>" & vbNewLine
		
		s = s & "<link type=""text/css"" rel=""stylesheet"" href=""" & Base_its & "/css/screen.css"">" & vbNewLine
		s = s & "<link type=""text/css"" rel=""stylesheet"" href=""" & Base_its & "/css/its.css"">" & vbNewLine

		if Include_about then
			s = s & "<link rel=""stylesheet"" href=""" & Base_its & "../common/about/css/about_screen.css"" type=""text/css"" media=""all"">" & vbNewLine
			s = s & "<script type=""text/javascript"" src=""" & Base_its & "/js/lib/jquery/jquery-1.2.3.pack.js""></script>" & vbNewLine
			s = s & "<script type=""text/javascript"" src=""" & Base_its & "../common/about/js/definitions.js""></script>" & vbNewLine
		end if

		if Include_cascade then
			s = s & "<script type=""text/javascript"" src=""" & Base_its & "/js/lib/prototype/prototype.js"" charset=""utf-8""></script>" & vbNewLine
			s = s & "<script type=""text/javascript"" src=""" & Base_its & "/js/cascade.js""></script>" & vbNewLine
		end if
		s = s & "</head>" & vbNewLine
		s = s & "<body " & BodyAttribs & ">" & vbNewLine


		s = s & "<div class=""mainlayout"">" & vbNewLine

		if Include_head then

			s = s & "<div class=""appheader"">" & vbNewLine
			s = s & "<div id=""logo""><a href=""../chart/default.asp""><img src=""/images/spacer.gif"" width=""158"" height=""80"" border=""0"" alt=""Altus Insite Retail Market"" /></a></div>" & vbNewLine
			s = s & "<div class=""navheader"">" & vbNewLine

			s = s & "<a href=""../chart/default.asp"">" & Lang("menu_home") & "</a> |" & vbNewLine
			s = s & "<a href=""../common/cross.asp?site=about"">" & Lang("menu_about") & "</a> |" & vbNewLine
			s = s & "<a href=""../common/contactform/default.asp"" target=""contact"" onclick=""window.open(this.href,this.target,'width=590,height=720,toolbar=no,scrollbars=yes,resizable=yes');return false;"" onkeypress=""this.click();"">" & Lang("menu_contact") & "</a> |" & vbNewLine

			if UserObj.IsLoggedIn then
				s = s & "<a href=""/logout.asp"">" & Lang("menu_logout") & "</a> |" & vbNewLine
			end if

			if Lang("lang")="fr-ca" then
				s = s & "<a href=""/eng/chart/default.asp"">" & Lang("lang_switch") & "</a>" & vbNewLine
			else
				s = s & "<a href=""/fre/chart/default.asp"">" & Lang("lang_switch") & "</a>" & vbNewLine
			end if

			s = s & "</div>" & vbNewLine
			s = s & "</div>" & vbNewLine

		end if

		s = s & "<table class=""searchpage"" border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">" & vbNewLine

		s = s & "<tr valign=""top"" align=""left"">" & vbNewLine

		if sidebar_area then

			s = s & "<td width=""159"">" & vbNewLine

			s = s & "<div id=""Sidelinks"">" & vbNewLine
			s = s & "<p>" & Lang("menu_label_producedby") & "</p>" & vbNewLine
			s = s & "<img src=""/images/side/altus.jpg"" alt=""" & Lang("alt_1") & """>" & vbNewLine
			s = s & "<hr style=""margin-bottom: 0px;"">" & vbNewLine
			s = s & "<div>" & vbNewLine
			s = s & "<ul>" & vbNewLine


			Set Input=new cInput
			toc_period=Input.period
			Set Input=nothing

			Set c = new cConverter

			toc_link="<a href=""../chart/toc.asp?period=" & toc_period & """>" & Lang("menu_currentqtrtoc") & "<br>" & c.NumToStr("LKP01_PERIOD_NAME", toc_period) & "</a>"
			toc_link="<li>" & toc_link & "</li>" & vbNewLine

			xls_files=Array("108", "208", "308", "408")
			i=toc_period-104
			xls_link=""
			if UserObj.IsLoggedIn and i>=0 and i<=ubound(xls_files) then
				xls_file=xls_files(i)
				xls_link="<a href=""/files/xls/IT" & xls_file & "-Excluded.xls"">" & Lang("menu_currentqtrxls") & "<br>" & c.NumToStr("LKP01_PERIOD_NAME", toc_period) & "</a>"
				xls_link="<li>" & xls_link & "</li>" & vbNewLine
			end if

			Set c = nothing


			s = s & toc_link
			s = s & xls_link

			s = s & "<li><a href=""../chart/hist.asp"">" & Lang("menu_historicalqtr") & "</a></li>" & vbNewLine


			s = s & "</ul>" & vbNewLine
			s = s & "<hr style=""margin-top: 15px;"">" & vbNewLine

			s = s & "<p>" & Lang("menu_label_additional") & "</p>" & vbNewLine
			s = s & "<ul>" & vbNewLine
			s = s & "<li><a class=""nopad"" href=""../common/cross.asp?site=office""><img src=""/images/side/office.jpg"" alt=""""></a>" & vbNewLine
			s = s & "<a href=""../common/cross.asp?site=office"">" & Lang("menu_additional_a") & "</a></li>" & vbNewLine
			s = s & "<li><a class=""nopad"" href=""../common/cross.asp?site=retail""><img src=""/images/side/retail.jpg"" alt=""""></a>" & vbNewLine
			s = s & "<a href=""../common/cross.asp?site=retail"">" & Lang("menu_additional_b") & "</a></li>" & vbNewLine
			s = s & "<li><a class=""nopad"" href=""../common/cross.asp?site=industrial""><img src=""/images/side/industrial.jpg"" alt=""""></a>" & vbNewLine
			s = s & "<a href=""../common/cross.asp?site=industrial"">" & Lang("menu_additional_c") & "</a></li>" & vbNewLine
			s = s & "<!--<li><a class=""nopad"" href=""../common/cross.asp?site=investment""><img src=""/images/side/investment.jpg"" alt=""""></a>" & vbNewLine
			s = s & "<a href=""../common/cross.asp?site=investment"">" & Lang("menu_additional_a") & "</a></li>-->" & vbNewLine
			s = s & "</ul>" & vbNewLine
			s = s & "<hr>" & vbNewLine

			s = s & "<a href=""http://www.space4lease.com/marketing/what_we_do.asp"" style=""padding: 0; margin: 0;"">" & vbNewLine
			s = s & "<img src=""/images/side/space4lease.gif"" width=""100"" height=""20"" alt=""Space4Lease""></a>" & vbNewLine
			s = s & "<a href=""http://www.space4lease.com/marketing/what_we_do.asp"" style=""padding: 0; margin: 0;"">" & Lang("menu_space4lease") & "</a>" & vbNewLine

			s = s & "<hr>" & vbNewLine
			s = s & "<p>" & Lang("menu_label_findmore") & "</p>" & vbNewLine
			s = s & "<ul>" & vbNewLine
			s = s & "<li><a href=""../common/contactform/default.asp"" target=""contact"" onclick=""window.open(this.href,this.target,'width=590,height=720,toolbar=no,scrollbars=yes,resizable=yes');return false;"" onkeypress=""this.click();"">" & Lang("menu_findmore_a") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""../common/cross.asp?site=about&amp;page=definitions_benchmark.php"">" & Lang("menu_findmore_b") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""../common/cross.asp?site=about&amp;page=marketdirectory_districtnodes.php"">" & Lang("menu_findmore_e") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""../common/cross.asp?site=about&amp;page=researchmethodology.php"">" & Lang("menu_findmore_c") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""../common/cross.asp?site=main&amp;page=copyright.php"">" & Lang("menu_findmore_d") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""../chart/freq.asp"">" & Lang("menu_qtrfrequency") & "</a></li>" & vbNewLine
			s = s & "</ul>" & vbNewLine
			s = s & "</div>" & vbNewLine
			s = s & "<hr style=""margin-top: 5px;"">" & vbNewLine

			if UserObj.AdminLevel=10 and UserObj.UserId=g_ITS_SUPER_USER then
				s = s & "<a href=""../chart/admin.asp"">" & Lang("menu_superuser") & "</a>" & vbNewLine
			end if

			s = s & "</div><!--/Sidelinks-->" & vbNewLine
			s = s & "</td>" & vbNewLine

		end if

		s = s & "<td class=""searchpagecontainer"">" & vbNewLine


		if sidebar_area then

			s = s & "<h1>" & Lang("head_default") & "</h1>" & vbNewLine

		else

			s = s & "<br>"
			s = s & "<div class=""logosplash"">"
			s = s & "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine
			s = s & "<tr valign=""top"">" & vbNewLine
			s = s & "<td align=""left"">" & vbNewLine

			if Len(UserObj.LogoFilename)>0 then
				s = s & "<img src=""" & Base_acs & "/images/logos/" & Server.HtmlEncode(UserObj.LogoFilename) & """ alt="""">"
			end if

			s = s & "</td>" & vbNewLine & vbNewLine
			s = s & "<td align=""left"">" & vbNewLine
			s = s & "<p>" & Lang("head_user_name") & " <strong>" & Server.HtmlEncode(UserObj.FirstName & " " & UserObj.LastName) & "&nbsp;</strong><br>" & vbNewLine
			s = s & Lang("head_user_email") & " <strong>" & Server.HtmlEncode(UserObj.Email) & "&nbsp;</strong><br>" & vbNewLine
			s = s & Lang("head_user_phone") & " <strong>" & Server.HtmlEncode(UserObj.Phone) & "&nbsp;</strong></p>" & vbNewLine
			s = s & "</td>" & vbNewLine
			s = s & "<td align=""right"">" & vbNewLine
			s = s & "<p>" & Lang("head_user_datetime") & " <strong>" & Now & "</strong></p>" & vbNewLine
			s = s & "</td>" & vbNewLine
			s = s & "</tr>" & vbNewLine
			s = s & "</table>" & vbNewLine
			s = s & "</div>" & vbNewLine

		end if


		s = s & "<h2 class=""subheadspecial"">" & Server.HtmlEncode(Subhead) & "</h2>" & vbNewLine

		if Include_its then
			s = s & "<div class=""its"">" & vbNewLine
		end if

		if nobuffer_content then
			Response.Write s
			ContentFunction()
			s=""
		else
			s = s & ContentFunction()
		end if

		if Include_its then
			s = s & "</div>" & vbNewLine
		end if




		s = s & "<div class=""footer"">" & vbNewLine

		if Include_pdf then

			s = s & "<p align=""center"">" & vbNewLine
			s = s & "<a href=""" & url & """><img src=""/images/footer/pdf.gif"" width=""112"" height=""45"" alt=""" & Lang("alt_2") & """ border=""0""></a>" & vbNewLine
			s = s & "</p>" & vbNewLine

		end if

		s = s & "<p class=""fineprint"">" & Lang("foot_copyright") & "</p>" & vbNewLine

		s = s & "</div>" & vbNewLine


		s = s & "</td>" & vbNewLine
		s = s & "</tr>" & vbNewLine

		s = s & "</table>" & vbNewLine

		s = s & "</div>" & vbNewLine


		s = s & "</body>" & vbNewLine
		s = s & "</html>" & vbNewLine

		if nobuffer then Response.Write s else DisplayEx=s

	End Function

End Class


Function GetLoginBox()

	s = ""

	s = s & "<br>" & vbNewLine
	s = s & "<form action=""/login.asp"" method=""post"" class=""objcenter searchform"" style=""width:50%;"">" & vbNewLine
	s = s & "<fieldset>" & vbNewLine
	s = s & "<legend>" & Lang("login") & "</legend>" & vbNewLine

	s = s & "<p>" & Lang("info_login") & "</p>" & vbNewLine

	s = s & "<p>" & vbNewLine
	s = s & "<label for=""uid"">" & Lang("userid") & "</label>" & vbNewLine
	s = s & "<input name=""uid"" id=""uid"" type=""text"" value="""" size=""40"" maxlength=""60""><br>" & vbNewLine
	s = s & "<label for=""pwd"">" & Lang("password") & "</label>" & vbNewLine
	s = s & "<input name=""pwd"" id=""pwd"" type=""password"" value="""" size=""40"" maxlength=""60"">" & vbNewLine
	s = s & "</p>" & vbNewLine

	s = s & "<input type=""hidden"" name=""url"" value=""" & Request.ServerVariables("SCRIPT_NAME") & """>" & vbNewLine

	's = s & "<div class=""textcenter"">" & vbNewLine
	s = s & "<input type=""submit"" value=""" & Lang("login") & """ class=""search"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../common/password.asp"">" & Lang("password_retrieval") & "</a>" & vbNewLine
	's = s & "</div>" & vbNewLine

	s = s & "<br><br><br><a href=""/login.asp?uid=" & g_ITS_GUEST_UID & "&amp;pwd=" & g_ITS_GUEST_PWD & "&amp;url=" & Request.ServerVariables("SCRIPT_NAME") & """>" & Lang("login2") & "</a>" & vbNewLine

	s = s & "</fieldset>" & vbNewLine

	s = s & "</form>" & vbNewLine

	GetLoginBox = s

End Function


Function GetInfoBox(str_error, str_back)

	s = ""
	s = s & "<br>&nbsp;" & vbNewLine
	s = s & "<h3 align=""center"">" & str_error & "</h3>" & vbNewLine
	s = s & "<script type=""text/javascript"">" & vbNewLine
	s = s & "<!-- to hide script contents from old browsers" & vbNewLine
	s = s & "document.write(""<p align=\""center\""><a href=\""javascript:history.go(-1)\"">" & str_back & "<\/a>"")" & vbNewLine
	s = s & "// end hiding contents from old browsers -->" & vbNewLine
	s = s & "</script>" & vbNewLine
	s = s & "<br>&nbsp;" & vbNewLine

	GetInfoBox = s

End Function

%>
