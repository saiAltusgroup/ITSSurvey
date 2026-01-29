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
		Include_reporthead=true 'Added for case when no sidebar but no report information required B.A.R.

		Base_acs="http://acs.altusinsite.com"
		
		'lacfre 2022-03 change for https
		'Base_its="http://" & Request.ServerVariables("SERVER_NAME")
		Base_its="https://" & Request.ServerVariables("SERVER_NAME")

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

	Public Function GetUrl(format, swaplang)
		urlvalues=Request.ServerVariables("QUERY_STRING")

		separator=""
		if Len(urlvalues)>0 then separator="&"

		for each thing in Request.Form

			If Request.Form(thing)<>"" Then
				urlvalues = urlvalues & separator & thing & "=" & Request.Form(thing)
				separator = "&"
			End If
		next

		url = Request.ServerVariables("SCRIPT_NAME")
		if (swaplang > 0) then
			if Lang("lang")="fr-ca" then
				url = Replace(url, "/fre/", "/eng/")
			else
				url = Replace(url, "/eng/", "/fre/")
			end if
		end if

		GetUrl = url & Server.HTMLEncode("?format=" & format & "&" & urlvalues)

	End Function

	Private Function DisplayEx(ContentArea, sidebar_area, nobuffer)

		DisplayEx=""
		nobuffer_content=nobuffer

		if Lang("lang")="fr-ca" then
			altus_url = g_SITE_MAIN & "fr/"
			home_url = "/fre/chart/"
		else
			altus_url = g_SITE_MAIN
			home_url = "/eng/chart/"
		end if

		if Protect and (not UserObj.IsLoggedIn) then

			sidebar_area=false
			Include_reporthead = false

			ContentArea = "GetLoginBox"
			nobuffer_content=false

			BodyAttribs=""

			Include_about=false
			Include_cascade=false
			Include_pdf=false

		end if

		Set ContentFunction = GetRef(ContentArea)

		' CONVERTED TO FUNcTION BY CRAIG September 23, 2014
		url = GetURL(1, 0)
		'urlvalues=Request.ServerVariables("QUERY_STRING")
                '
		'separator=""
		'if Len(urlvalues)>0 then separator="&"
                '
		'for each thing in Request.Form
                '
		'	If Request.Form(thing)<>"" Then
		'		urlvalues = urlvalues & separator & thing & "=" & Request.Form(thing)
		'		separator = "&"
		'	End If
		'next
                '
		'url=Request.ServerVariables("SCRIPT_NAME") & Server.HTMLEncode("?format=1&" & urlvalues)


		s = ""

		s = s & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbNewLine
		s = s & "<html lang=""" & Lang("lang") & """>" & vbNewLine
		s = s & "<head>" & vbNewLine

		s = s & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbNewLine
		s = s & "<title>" & Title & "</title>" & vbNewLine

		s = s & "<style media=""all"" type=""text/css"">" & vbNewLine
		s = s & "<!--" & vbNewLine
		s = s & " @import url('" & Base_its & "/include/presentation/css/style.css');" & vbNewLine
		s = s & "-->" & vbNewLine
		s = s & "</style>" & vbNewLine
		s = s & "<!--[if lte IE 7]>" & vbNewLine
		s = s & "<style media=""all"" type=""text/css"">" & vbNewLine
		s = s & "	@import url('" & Base_its & "/include/presentation/css/ie.css');" & vbNewLine
		s = s & "</style>" & vbNewLine
		s = s & "<![endif]-->" & vbNewLine
		s = s & "<!--[if lt IE 7]>" & vbNewLine
		s = s & "<style media=""all"" type=""text/css"">" & vbNewLine
		s = s & "	@import url('" & Base_its & "/include/presentation/css/ie6.css');" & vbNewLine
		s = s & "</style>" & vbNewLine
		s = s & "<![endif]-->" & vbNewLine

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
			s = s & "<script type=""text/javascript"">setTimeout(function(){updateMarket();}, 2000);</script>" & vbNewLine
		end if

		s = s & "</head>" & vbNewLine
		s = s & "<body class=""investment"" " & BodyAttribs & ">" & vbNewLine
		s = s & "<a class=""access"" href=""#Content"">Skip to content</a>" & vbNewLine

		s = s & "<div id=""Main"">" & vbNewLine

    s = s & "<div id=""topbar"">" & vbNewLine
			s = s & "<ul class=""menu"">" & vbNewLine
			s = s & "	<li class=""first""><a href=""" & home_url & """ class=""first"" title=""" & Lang("menu_home") & """>" & Lang("menu_home") & "</a></li>" & vbNewLine
			s = s & "	<li><a href=""" & altus_url & "index.php"" title=""" & Lang("menu_search") & """>" & Lang("menu_search") & "</a>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "	<li><a href=""" & altus_url & "index.php?page=savedreport_home&amp;action=list"" title=""My Reports"">" & Lang("menu_reports") & "</a>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "	<li><a href=""" & altus_url & "index.php?page=contact"" title=""Contact"">" & Lang("menu_contact") & "</a></li>" & vbNewLine
			s = s & "	<li class=""last"">" & vbNewLine
			s = s & "	<a href=""" & GetUrl(0, 1) & """ title=""" & Lang("lang_switch") & """>" & Lang("lang_switch") & "</a>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "</ul>" & vbNewLine
    s = s & "</div>" & vbNewLine

		's = s & "<div class=""mainlayout"">" & vbNewLine

'need to get language translations

		if Include_head then

			s = s & "	<div id=""Head"" class=""clearfix"">" & vbNewLine
'			s = s & "		<a href=""" & altus_url & "index.php"" title=""Altus InSite - Canada's Leading Real Estate Information Provider""><img id=""Logo"" class=""altus"" src=""/images/tmpl/logo.png"" width=""118"" height=""48"" alt=""Altus InSite"" /></a>" & vbNewLine
            if Lang("lang")="fr-ca" then
			    s = s & "		<div id=""SiteLogo"" ><a href=""" & altus_url & "index.php""><img style=""margin-top:11px;margin-left: 19px;"" src=""/images/tmpl/AltusGroup_AltusInsite_NavBar_Logo_FR.png"" alt=""Altus InSite - Canada's Leading Real Estate Information Provider"" /></a></div>" & vbNewLine
			else
			    s = s & "		<div id=""SiteLogo"" ><a href=""" & altus_url & "index.php""><img style=""margin-top:11px;margin-left: 13px;"" src=""/images/tmpl/AltusGroup_AltusInsite_NavBar_Logo.png"" alt=""Canada's Leading Real Estate Information Provider"" /></a></div>" & vbNewLine
			end if
			s = s & "	</div><!--/Head-->" & vbNewLine
			s = s & "<div id=""Holder"" class=""clearfix"">" & vbNewLine
			if sidebar_area then
					s = s & "<div id=""Panel1"">" & vbNewLine
			end if

			's = s & "<ul id=""TopMenu"">" & vbNewLine
			's = s & "	<li class=""first""><a href=""" & g_SITE_MAIN & "index.php"" class=""first"" title=""Home"">" & Lang("menu_home") & "</a></li>" & vbNewLine
			's = s & "	<li class=""selected""><a href=""" & g_SITE_MAIN & "index.php"" title=""Search""><em>" & Lang("menu_search") & "</em></a>" & vbNewLine
			's = s & "	</li>" & vbNewLine
			's = s & "	<li><a href=""" & g_SITE_MAIN & "index.php?page=myreports"" title=""My Reports"">" & Lang("menu_reports") & "</a>" & vbNewLine
			's = s & "	</li>" & vbNewLine
			's = s & "	<li><a href=""" & g_SITE_MAIN & "index.php?page=contact"" title=""Contact"">" & Lang("menu_contact") & "</a></li>" & vbNewLine
			's = s & "	<li class=""last"">" & vbNewLine
			'if Lang("lang")="fr-ca" then
				's = s & "<a href=""/eng/chart/default.asp"" title=""English"">" & Lang("lang_switch") & "</a>" & vbNewLine
			'else
				's = s & "<a href=""/fre/chart/default.asp"" title=""Fran&ccedil;ais"">" & Lang("lang_switch") & "</a>" & vbNewLine
			'end if
			's = s & "	</li>" & vbNewLine
			's = s & "</ul>" & vbNewLine

		end if



		s = s & "<div id=""Content"">" & vbNewLine

		if(sidebar_area and Include_reporthead) then

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
			margin_style_h1 = ""
			margin_style_h2 = ""

		elseif NOT sidebar_area then
			margin_style_h1 = " style=""margin-left: 1em; margin-top: 1em;"""
			margin_style_h2 = " style=""margin-left: 1.3em;"""
		end if

		s = s & "<h1" & margin_style_h1 & ">" & Lang("head_default") & "</h1>" & vbNewLine
		s = s & "<h2" & margin_style_h2 & " class=""subheadspecial"">" & Server.HtmlEncode(Subhead) & "</h2>" & vbNewLine

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


		if sidebar_area then

			s = s & "<div id=""Panel2"">" & vbNewLine


			s = s & "	<div id=""UserArea"">" & vbNewLine
			s = s & "		<p>" & vbNewLine
			s = s & "			" & Lang("menu_loggedin") & ":<br />" & vbNewLine
			s = s & "			<span class=""clientname"">" & UserObj.FirstName & " " & UserObj.LastName & "</span><br />" & vbNewLine
			
			'lacfre 2021-10-18
			's = s & "			<a href=""" & altus_url & "index.php?page=register"">" & Lang("menu_myaccount") & "</a> | <a href=""/logout.asp"">" & Lang("menu_logout") & "</a>" & vbNewLine
			s = s & "			<a href=""" & altus_url & "index.php?page=edit_profile"">" & Lang("menu_myaccount") & "</a> | <a href=""/logout.asp"">" & Lang("menu_logout") & "</a>" & vbNewLine
			
			s = s & "		</p>" & vbNewLine
			s = s & "	</div>" & vbNewLine

			s = s & "<div id=""Sidelinks"">" & vbNewLine
			s = s & "	<ul>" & vbNewLine


			Set Input=new cInput
			toc_period=Input.period
			Set Input=nothing

			Set c = new cConverter

			toc_link="<a href=""../chart/toc.asp?period=" & toc_period & """>" & Lang("menu_currentqtrtoc") &  c.NumToStr("LKP01_PERIOD_NAME", toc_period) & "</a>"
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

			s = s & "<li><a href=""../chart/default.asp"">" & Lang("menu_currentqtr") & "</a></li>" & vbNewLine
			s = s & toc_link
			s = s & xls_link

			s = s & "<li><a href=""../chart/hist.asp"">" & Lang("menu_historicalqtr") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""../chart/freq.asp"">" & Lang("menu_qtrfrequency") & "</a></li>" & vbNewLine

			if Lang("lang")="fr-ca" then
				s = s & "<li><a href=""http://altusinsite.com/fr/index.php?page=its/definitions_terms"" target=""_blank"">" & Lang("menu_defterm") & "</a></li>" & vbNewLine
				s = s & "<li><a href=""http://altusinsite.com/fr/index.php?page=its/benchmark_properties"" target=""_blank"">" & Lang("menu_benchprop") & "</a></li>" & vbNewLine
        	else
			s = s & "<li><a href=""http://altusinsite.com/en/index.php?page=its/definitions_terms"" target=""_blank"">" & Lang("menu_defterm") & "</a></li>" & vbNewLine
			s = s & "<li><a href=""http://altusinsite.com/en/index.php?page=its/benchmark_properties"" target=""_blank"">" & Lang("menu_benchprop") & "</a></li>" & vbNewLine
            end if

			s = s & "</ul>" & vbNewLine
s = s & "</div><!--/Sidelinks-->" & vbNewLine
			's = s & "	<div id=""AdditionalDataPerspectives"">" & vbNewLine
			's = s & "		<ul>" & vbNewLine
			's = s & "			<li>" & vbNewLine
			's = s & "				<a class=""nopad"" href=""../common/cross.asp?site=office""><img src=""/image/en/side/logo_office_small.jpg"" alt="""" /></a>" & vbNewLine
			's = s & "				<a href=""../common/cross.asp?site=office"">" & Lang("menu_additional_a") & "</a>" & vbNewLine
			's = s & "			</li>" & vbNewLine
			's = s & "			<li>" & vbNewLine
			's = s & "				<a class=""nopad"" href=""../common/cross.asp?site=industrial""><img src=""/image/en/side/logo_industrial_small.jpg"" alt="""" /></a>" & vbNewLine
			's = s & "				<a href=""../common/cross.asp?site=industrial"">" & Lang("menu_additional_c") & "</a>" & vbNewLine
			's = s & "			</li>" & vbNewLine
'			s = s & "			<li>" & vbNewLine
'			s = s & "				<a class=""nopad"" href=""../common/cross.asp?site=retail""><img src=""/image/en/side/logo_retail_small.jpg"" alt="""" /></a>" & vbNewLine
'			s = s & "				<a href=""../common/cross.asp?site=retail"">" & Lang("menu_additional_b") & "</a>" & vbNewLine
'			s = s & "			</li>" & vbNewLine

			's = s & "		</ul>" & vbNewLine
			's = s & "	</div>" & vbNewLine

			if UserObj.AdminLevel=100 and UserObj.UserId=g_ITS_SUPER_USER then
				s = s & "<a href=""../chart/admin.asp"">" & Lang("menu_superuser") & "</a>" & vbNewLine
			end if

			s = s & "</div><!--/Panel2-->" & vbNewLine


		end if

		if Include_pdf then
			s = s & "<div class=""footer"">" & vbNewLine
			s = s & "	<p align=""center"">" & vbNewLine
			s = s & "<p align=""center"" style=""color:red;""><b>Use<br/>CTRL + P<br/>to print a PDF</b></p>"
			's = s & "		<a href=""" & url & """><img src=""/images/footer/pdf.gif"" width=""35"" height=""35"" alt=""" & Lang("alt_2") & """ border=""0""><br />" & Lang("alt_2") & "</a>" & vbNewLine
			s = s & "	</p>" & vbNewLine
			s = s & "</div>" & vbNewLine
		end if



		s = s & "</div><!--/Content-->" & vbNewLine
		if sidebar_area then
				s = s & "</div><!--/Panel1-->" & vbNewLine
		end if
		s = s & "</div><!--/Holder-->" & vbNewLine
		s = s & "</div><!--/Main-->" & vbNewLine
		s = s & "<div id=""Foot"">" & vbNewLine

			s = s & "		<ul class=""menu"">" & vbNewLine
			s = s & "			<li><a href=""" & altus_url & "index.php?page=contact"" title=""" & Lang("menu_contact") & """>" & Lang("menu_contact") & "</a></li>" & vbNewLine
			s = s & "			<li><a href=""http://" & Lang("menu_altusgroup_url") & """ title=""" & Lang("menu_altusgroup") & " (http://" & Lang("menu_altusgroup_url") & ")"" target=""_blank"">" & Lang("menu_altusgroup") & "</a></li>" & vbNewLine
      s = s & "			<li><a href=""" & altus_url & "index.php?page=privacy"" title=""" & Lang("menu_privacy") & """>" & Lang("menu_privacy") & "</a></li>" & vbNewLine
      s = s & "			<li class=""last""><a href=""" & altus_url & "index.php?page=copyright"" title=""" & Lang("menu_copyright") & """>" & Lang("menu_copyright") & "</a></li>" & vbNewLine
			s = s & "		</ul>" & vbNewLine

      s = s & "<address>" & Lang("foot_copyright") & "</address>" & vbNewLine

		s = s & "</div><!--/Foot-->" & vbNewLine

		s = s & "<script type=""text/javascript""><!--" & vbNewLine
		s = s & "var gaJsHost = ((""https:"" == document.location.protocol) ? ""https://ssl."" : ""http://www."");" & vbNewLine
		s = s & "document.write(unescape(""%3Cscript src='"" + gaJsHost + ""google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E""));" & vbNewLine
		s = s & "</script>" & vbNewLine
		s = s & "<script type=""text/javascript"">" & vbNewLine
		s = s & "try{" & vbNewLine
		s = s & "var pageTracker = _gat._getTracker(""UA-12306099-4"");" & vbNewLine
		s = s & "pageTracker._trackPageview();" & vbNewLine
		s = s & "} catch(err) {} " & vbNewLine
		s = s & "// --></script>" & vbNewLine

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

	s = s & Lang("info_login") & vbNewLine

	If (Request.QueryString("failed") = "1") Then s = s & "<p class=""objcenter"" style=""width: 244px; color: #f00;"">" & Lang("info_login_incorrect") & "</p>"

	s = s & "<p>" & vbNewLine
	s = s & "<label for=""uid"">" & Lang("userid") & "</label>" & vbNewLine
	s = s & "<input name=""uid"" id=""uid"" type=""text"" value="""" size=""40""><br>" & vbNewLine
	s = s & "<label for=""pwd"">" & Lang("password") & "</label>" & vbNewLine
	s = s & "<input name=""pwd"" id=""pwd"" type=""password"" value="""" size=""40"">" & vbNewLine
	s = s & "</p>" & vbNewLine

	s = s & "<input type=""hidden"" name=""url"" value=""" & Request.ServerVariables("SCRIPT_NAME") & """>" & vbNewLine

	's = s & "<div class=""textcenter"">" & vbNewLine
	s = s & "<input type=""submit"" value=""" & Lang("login") & """ class=""search"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""../common/password.asp"">" & Lang("password_retrieval") & "</a>" & vbNewLine
	's = s & "</div>" & vbNewLine

	's = s & "<br><br><br><a href=""/login.asp?uid=" & g_ITS_GUEST_UID & "&amp;pwd=" & g_ITS_GUEST_PWD & "&amp;url=" & Request.ServerVariables("SCRIPT_NAME") & """>" & Lang("login2") & "</a>" & vbNewLine

	s = s & "<br /><br />" & Lang("info_login2") & vbNewLine
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
