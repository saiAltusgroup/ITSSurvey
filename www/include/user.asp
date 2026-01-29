<%

' DK, Apr 01, 2008 - version 1.0
' DK, Dec 03, 2010 - changed code to use the new AltusInsite DB for user authentication

'lacfre 2021-10-18
Const SALT_KEY_PASSWORD="$1$zMALl5Bh"

Class cUser
	Public IsLoggedIn
	Public UserId

	Public FirstName
	Public LastName
	Public Phone
	Public Email
	Public LogoFilename
	Public AdminLevel

	' Constructor '
	Private Sub Class_Initialize()

		IsLoggedIn=false
		UserId=""

		FirstName=""
		LastName=""
		Phone=""
		Email=""
		LogoFilename=""
		AdminLevel=0

		if session("rid") = "3A2E17CA24FA6404BE21" and len(session("uid"))>0 Then

			IsLoggedIn=true
			UserId=session("uid")

			FirstName=session("f1")
			LastName=session("f2")
			Phone=session("f3")
			Email=session("f4")
			LogoFilename=session("f5")
			AdminLevel=session("f6")

		end if

	End Sub

	Public Function Login(user_id, user_password)
		Login=false

		UserId=user_id
		UserPassword=user_password
		

		user_id = Replace(UserId, "'", "''")
		user_password = Replace(UserPassword, "'", "''")

		'sql = "select bold_usertable.usertableid, bold_usertable.userid, bold_usertable.password, bold_usertable.firstname, bold_usertable.lastname, bold_usertable.phone, bold_usertable.email, bold_usertable.adminlevel, bold_usercompany.logofilename from bold_usertable left outer join bold_usercompany on bold_usertable.usercompanyid = bold_usercompany.usercompanyid where userid='" & user_id & "' and password='" & user_password & "'"


		'lacfre 2021-10-18
		encryptPassword = GetEncryptPassword(user_password)
		user_password = Replace(encryptPassword, "'", "''")
		

		'lacfre 2021-10-18
		'sql = "select Username, FirstName, LastName, P.Phone, P.Email, O.LogoFilename, AdminLevel from [User] U join Person P on U.Person_ID=P.ID left outer join Organization O on P.Organization_ID=O.ID where Username='" & user_id & "' and Password='" & user_password & "'"
		sql = "exec [dbo].[usp_User_Login] '" + user_id + "','" + user_password + "'"
	
		'Response.Write sql
		


		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql, acs_conn, 3, 1

		'lacfre 2021-10-18
		'If (rs.RecordCount = 1) Then
		If(not rs.EOF) Then
			session("rid") = "3A2E17CA24FA6404BE21"
			session("uid") = UserId

			session("f1") = rs("firstname")
			session("f2") = rs("lastname")
			session("f3") = rs("phone")
			session("f4") = rs("email")
			session("f5") = rs("logofilename")
			session("f6") = rs("adminlevel")

			Login=true

		End If

		rs.close

		insert_firstname = Replace(session("f1"), "'", "''")
		insert_lastname = Replace(session("f2"), "'", "''")
		If (NOT IsNull(session("f3"))) Then insert_phone = Replace(session("f3"), "'", "''")
		insert_email = Replace(session("f4"), "'", "''")
		insert_datein = Replace((Month(Date) & "-" & Day(Date) & "-" & Year(Date) & " " & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)), "'", "''")
		insert_ip = Replace(Request.ServerVariables("REMOTE_ADDR"), "'", "''")
		insert_host = Replace(Request.ServerVariables("REMOTE_HOST"), "'", "''")
		insert_acceptencoding = Replace(Request.ServerVariables("HTTP_ACCEPT_ENCODING"), "'", "''")
		insert_useragent = Replace(Request.ServerVariables("HTTP_USER_AGENT"), "'", "''")
		insert_loginname = Replace(session("uid"), "'", "''")

		sql = "insert into zloguser (name, email, phone, datein, ip, host, acceptencoding, useragent, Login_Name, First_Name, Last_Name) values ('" & insert_firstname & " " & insert_lastname & "', '" & insert_email & "', '" & insert_phone & "', '" & insert_datein & "', '" & insert_ip & "', '" & insert_host & "', '" & insert_acceptencoding & "', '" & insert_useragent & "', '" & insert_loginname & "', '" & insert_firstname & "', '" & insert_lastname & "')"
		rs.open sql, its_conn, 3, 1

		Set rs=Nothing

	End Function


	'lacfre 2020-10-18
	Public Function GetEncryptPassword(clearPassword)
		encryptPassword = ""
	
		sql = "select dbo.EncryptPassword('" + clearPassword + "','" + SALT_KEY_PASSWORD + "')"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql, acs_conn, 3, 1
		encryptPassword = rs(0)
		rs.close

		GetEncryptPassword = encryptPassword
	End Function



	Public Function GetPassword()

		GetPassword = ""

		user_id = Replace(UserId, "'", "''")

		'sql = "select bold_usertable.password from bold_usertable where userid='" & user_id & "'"

		sql = "select Password from [User] where Username='" & user_id & "'"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql, acs_conn, 3, 1

		If (rs.RecordCount = 1) Then
			GetPassword = rs("password")
		End If

		rs.close
		Set rs=Nothing

	End Function


	Public Function HasAccessToAllMarketsForReport(report_code)

		HasAccessToAllMarketsForReport = false

		'return_array_counter = 0
		'return_array = array(0)
		'return_array(0) = 0

		'GetPermittedMarketsForReport = return_array
		
		If Len(UserId)<1 or Len(report_code)<1 Then Exit Function

		'current_quarter = 1

		'productID = 4

		'current_month = Month(Date())
		'current_year = Year(Date())

		'Select Case current_month
		'	Case 1, 2, 3
		'		current_quarter = 1
		'	Case 4, 5, 6
		'		current_quarter = 2
		'	Case 7, 8, 9
		'		current_quarter = 3
		'	Case 10, 11, 12
		'		current_quarter = 4
		'End Select

		' Query based on UserId and report code

		user_id = Replace(UserId, "'", "''")

		'sql = "SELECT DISTINCT bra.access_marketID FROM ((((bold_usertable bu LEFT OUTER JOIN bold_user_group bug on bu.usertableid = bug.ug_userID) LEFT OUTER JOIN bold_report_access bra ON bug.ug_groupID = bra.access_groupID) LEFT OUTER JOIN bold_report_access_expiry brae ON bug.ug_groupID = brae.expiry_groupID) INNER JOIN bold_report br on bra.access_reportID = br.report_ID) WHERE bu.userid = '" & user_id & "' and br.report_code = '" & report_code & "' and bra.access_productID = " & productID & " AND brae.expiry_qtr" & current_quarter & " = 'yes' AND (brae.expiry_year >= " & current_year & " OR brae.expiry_year IS NULL OR brae.expiry_year = '')"

		sql = "select GeozoneIDArray from [User] U join Map_UserGroup_User MUU on U.ID = MUU.User_ID join Map_UserGroup_AccessControlRule MUA on MUU.UserGroup_ID = MUA.UserGroup_ID join AccessControlRule A ON MUA.AccessControlRule_ID = A.ID where A.ReportTypeArray like '%" & g_ITS_REPORT_CODE & "%' and U.Username = '" & user_id & "'"

		'Response.Write sql & "<br>"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql, acs_conn, 3, 1
		record_count=rs.RecordCount
		rs.close
		Set rs=Nothing

		'Response.Write record_count

		If (record_count >= 1) Then HasAccessToAllMarketsForReport = true

	End Function

	  
	
	

	Public Function SendPassword(email)

		SendPassword = false

		email=Trim(email)
		If Len(email)<1 Then Exit Function

		email = Replace(email, "'", "''")

		'sql = "select userid, password from bold_usertable where email = '" & email & "'"

		'lacfre 2021-10-18
		'sql = "select Username as UserID, Password from [User] U join Person P on U.Person_ID=P.ID where Email='" & email & "'"
		sql = "exec [dbo].[usp_User_Get_ByEmail] '"  & email & "'"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.open sql, acs_conn, 3, 1


		'lacfre 2021-10-18
		'If (rs.RecordCount = 1) Then
		If(not rs.EOF) Then

			'lacfre 2021-10-18
			'password = rs("password")
			password = Mid(StrShuffle("!@#$%*&"),1,2) & Mid(StrShuffle("abcdefghijklmnpqrstuwxyz"),1,4) &_ 
					   Mid(StrShuffle("ABCDEFGHJKLMNPQRSTUWXYZ"),1,4) & Mid(StrShuffle("0123456789"),1,2)
			password = StrShuffle(password)
			user_name = rs("username")			

			user_id = rs("User_ID")
			

			'email_from = "info@altusinsite.com"
			email_from = "info_insite@aws-ses.altusgroup.com" 

			subject = "Altus InSite Password Retrieval"
			'message = "Here is your Altus InSite Registration information."& VBNewLine & VBNewLine &_
			message =  "Here is a temporary password for 24 hours. You must change your password in the edit profile section of the altusinsite website." & VBNewLine & VBNewLine &_ 
						"User ID: "& user_name & VBNewLine &_
						"Password: "& password & VBNewLine & VBNewLine &_
						"To edit profile, go to " & g_SITE_MAIN & "?page=edit_profile" & VBNewLine & VBNewLine &_ 
						"If you have problems with your login, please contact us at useradmin@altusinsite.com."
						'"To log in, go to http://its.altusinsite.com"& VBNewLine

			Set mailer = CreateObject("CDO.Message")

			If isobject(mailer)=true Then

				mailer.From = email_from
				mailer.Subject = subject
				mailer.To = email
				mailer.TextBody = message
				mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = g_SITE_SMTP_IPADDR
				mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = g_SITE_SMTP_PORT

				'lacfre 2021-10-18
				mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "AKIA5EJU6MB4PEXI6R4R"
				mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "BK4D3L7MsfBneoyGU21ZKOCxchAbhXWrvvWZySkRGCdd"
				mailer.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1

				mailer.Configuration.Fields.Update
				mailer.Send

				encryptPassword = GetEncryptPassword(password)
				sql = "UPDATE dbo.[user] set PasswordTmp='" & Replace(encryptPassword,"'", "''") & "',PasswordTmpExpiration=DATEADD(HOUR, 24, getdate()) WHERE ID=" & user_id
				Set rs2 = Server.CreateObject("ADODB.Recordset")
				rs2.open sql, acs_conn, 3, 1

				SendPassword = true

			End If

			set mailer = nothing

		End If

		rs.close
		Set rs=Nothing

	End Function

	'lacfre 2021-10-18
	Function StrShuffle(str)
		rtn = ""
        i = 0

        Do while len(str) > 0
            i = Int((len(str))*Rnd+1)
			carac = Mid(str, i, 1)
            rtn = rtn + carac
            str = Replace(str,carac,"",1)
        Loop

        StrShuffle = rtn

	End Function

End Class

%>
