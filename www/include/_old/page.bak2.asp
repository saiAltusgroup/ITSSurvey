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
		end if

		s = s & "</head>" & vbNewLine
		s = s & "<body class=""investment"" " & BodyAttribs & ">" & vbNewLine
		s = s & "<a class=""access"" href=""#Content"">Skip to content</a>" & vbNewLine
		s = s & "<div id=""Main"">" & vbNewLine


		's = s & "<div class=""mainlayout"">" & vbNewLine

'need to get language translations

		if Include_head then

			s = s & "	<div id=""Head"" class=""clearfix"">" & vbNewLine
			s = s & "		<a href=""../common/cross.asp?site=main""><img id=""Logo"" class=""altus"" src=""/image/en/head/logo_altus.jpg"" alt=""Altus InSite"" /></a>" & vbNewLine
			s = s & "		<div id=""SiteLogo"" ><a href=""../chart/default.asp""><img src=""/image/en/head/logo_investment.jpg"" alt=""Investment Trends Survey"" /></a></div>" & vbNewLine
			s = s & "	</div><!--/Head-->" & vbNewLine
			s = s & "<div id=""Holder"" class=""clearfix"">" & vbNewLine
			if sidebar_area then
					s = s & "<div id=""Panel1"">" & vbNewLine
			end if

			s = s & "<ul id=""TopMenu"">" & vbNewLine
			s = s & "	<li class=""first""><a href=""../common/cross.asp?site=main"" class=""first"" title=""Home"">" & Lang("menu_home") & "</a></li>" & vbNewLine
			s = s & "	<li class=""selected""><a href=""../common/cross.asp?site=main&amp;page=/search/index.php"" title=""Search""><em>" & Lang("menu_search") & "</em></a>" & vbNewLine
			s = s & "		<div class=""level2"">" & vbNewLine
			s = s & "			<ul>" & vbNewLine
			s = s & "				<li class=""first""><a href=""../common/cross.asp?site=office"" class=""first"">" & Lang("menu_office") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=industrial"">" & Lang("menu_industrial") & "</a></li>" & vbNewLine
			s = s & "				<li class=""last selected""><a href=""../common/cross.asp?site=investment"" class=""last""><em>" & Lang("menu_ITS") & "</em></a></li>" & vbNewLine
			s = s & "			</ul>" & vbNewLine
			s = s & "		</div>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "	<li><a href=""../common/cross.asp?site=myreports"">" & Lang("menu_reports") & "</a>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "	<li><a href=""../common/cross.asp?site=main&amp;page=/freetools/index.php"">" & Lang("menu_freetools") & "</a>" & vbNewLine
			s = s & "		<div class=""level2"">" & vbNewLine
			s = s & "			<ul>" & vbNewLine
			s = s & "				<li class=""first""><a href=""http://search.altusinsite.com/main/search.asp"" class=""first"">" & Lang("menu_search") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=myreports&amp;page=/add/add_step1.asp"" class=""first"">" & Lang("menu_addlisting") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=myreports&amp;page=/main/request_link.asp"">" & Lang("menu_updateproperties") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=myreports&amp;page=/main/resource_centre/default.asp"">" & Lang("menu_tenantresourcecentre") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/freetools/research_report.php"">" & Lang("menu_researchreports") & "</a></li>" & vbNewLine
			s = s & "				<li class=""last""><a href=""../common/cross.asp?site=main&amp;page=/freetools/submit_transactions.php"" class=""last"">" & Lang("menu_submittransactions") & "</a></li>" & vbNewLine
			s = s & "			</ul>" & vbNewLine
			s = s & "		</div>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "	<li><a href=""../common/cross.asp?site=main&amp;page=/about/index.php"" class=""selected"" title=""About Altus InSite"">" & Lang("menu_about") & "</a>" & vbNewLine
			s = s & "		<div class=""level2"">" & vbNewLine
			s = s & "			<ul>" & vbNewLine
			s = s & "				<li class=""first""><a href=""../common/cross.asp?site=main&amp;page=/about/whatwedo.php"" class=""first"" >" & Lang("menu_whatwedo") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/subscription.php"" title=""Subscription Options"">" & Lang("menu_subscription") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/definitions_office.php"" title=""Office Definitions"">" & Lang("menu_officedefinitions") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/marketdirectory_districtnodes.php"" title=""Market Directory District Nodes"">" & Lang("menu_officegeography") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/definitions_industrial.php"" title=""Industrial Definitions"">" & Lang("menu_industrialdefinitions") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/definitions_benchmark.php"" title=""Benchmark Properties"">" & Lang("menu_benchmarkproperties") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/researchmethodology.php"" title=""Research Methodology"">" & Lang("menu_researchmethodology") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/definitions_barometer.php"" title=""Barometer Definitions"">" & Lang("menu_barometerdefinitions") & "</a></li>" & vbNewLine
			s = s & "				<li><a href=""../common/cross.asp?site=main&amp;page=/about/definitions_retail.php"" class=""last"" title=""Retail Definitions"">" & Lang("menu_retaildefinitions") & "</a></li>" & vbNewLine
			s = s & "				<li class=""last""><a href=""../common/cross.asp?site=main&amp;page=/copyright.php"" title=""Copyright Details"">" & Lang("menu_copyright") & "</a></li>" & vbNewLine
			s = s & "			</ul>" & vbNewLine
			s = s & "		</div>" & vbNewLine
			s = s & "	</li>" & vbNewLine
			s = s & "	<li class=""thickbox""><a href=""../common/cross.asp?site=main&amp;page=/contactform/contactus.php?KeepThis=true&amp;TB_iframe=true&amp;height=650&amp;width=590"" class=""thickbox last"" target=""Contact"">" & Lang("menu_contact") & "</a></li>" & vbNewLine
			s = s & "	<li class=""last"">" & vbNewLine
			if Lang("lang")="fr-ca" then
				s = s & "<a href=""/eng/chart/default.asp"">" & Lang("lang_switch") & "</a>" & vbNewLine
			else
				s = s & "<a href=""/fre/chart/default.asp"">" & Lang("lang_switch") & "</a>" & vbNewLine
			end if
			s = s & "	</li>" & vbNewLine
			s = s & "</ul>" & vbNewLine

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
			margin_style_h1 = " style=""margin-left: 1em;"""
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
			s = s & "			<a href=""../common/cross.asp?site=main&amp;page=register.php"">" & Lang("menu_myaccount") & "</a> | <a href=""/logout.asp"">" & Lang("menu_logout") & "</a>" & vbNewLine
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
			s = s & "</ul>" & vbNewLine
s = s & "</div><!--/Sidelinks-->" & vbNewLine
			s = s & "	<div id=""AdditionalDataPerspectives"">" & vbNewLine
			s = s & "		<ul>" & vbNewLine
			s = s & "			<li>" & vbNewLine
			s = s & "				<a class=""nopad"" href=""../common/cross.asp?site=office""><img src=""/image/en/side/logo_office_small.jpg"" alt="""" /></a>" & vbNewLine
			s = s & "				<a href=""../common/cross.asp?site=office"">" & Lang("menu_additional_a") & "</a>" & vbNewLine
			s = s & "			</li>" & vbNewLine
			s = s & "			<li>" & vbNewLine
			s = s & "				<a class=""nopad"" href=""../common/cross.asp?site=industrial""><img src=""/image/en/side/logo_industrial_small.jpg"" alt="""" /></a>" & vbNewLine
			s = s & "				<a href=""../common/cross.asp?site=industrial"">" & Lang("menu_additional_c") & "</a>" & vbNewLine
			s = s & "			</li>" & vbNewLine
'			s = s & "			<li>" & vbNewLine
'			s = s & "				<a class=""nopad"" href=""../common/cross.asp?site=retail""><img src=""/image/en/side/logo_retail_small.jpg"" alt="""" /></a>" & vbNewLine
'			s = s & "				<a href=""../common/cross.asp?site=retail"">" & Lang("menu_additional_b") & "</a>" & vbNewLine
'			s = s & "			</li>" & vbNewLine

			s = s & "		</ul>" & vbNewLine
			s = s & "	</div>" & vbNewLine

			if UserObj.AdminLevel=10 and UserObj.UserId=g_ITS_SUPER_USER then
				s = s & "<a href=""../chart/admin.asp"">" & Lang("menu_superuser") & "</a>" & vbNewLine
			end if

			s = s & "</div><!--/Panel2-->" & vbNewLine


		end if

		if Include_pdf then
			s = s & "<div class=""footer"">" & vbNewLine
			s = s & "	<p align=""center"">" & vbNewLine
			s = s & "		<a href=""" & url & """><img src=""/images/footer/pdf.gif"" width=""112"" height=""45"" alt=""" & Lang("alt_2") & """ border=""0""></a>" & vbNewLine
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

	s = s & "<p>" & Lang("info_login") & "</p>" & vbNewLine

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
