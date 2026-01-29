<%
	'Session.Timeout=20

	if isObject(session("its_conn")) then
		set its_conn=session("its_conn")
	else
		set its_conn=server.createobject("ADODB.Connection")
		session("its_conn")=its_conn
		'its_conn.open "Provider=SQLOLEDB.1; Data Source=10.0.1.30; Initial Catalog=ITSLive; User ID=itslive; Password=dak!W4yeQ;"
		its_conn.open "Provider=SQLOLEDB.1; Data Source=sql-qa-use1.altusinsite.com; Initial Catalog=ITSLive; User ID=ITSLive; Password=ITSLive;"
		'its_conn.open "Provider=SQLOLEDB.1; Data Source=192.168.0.2; Initial Catalog=ITSLive; User ID=ITSLive; Password=ITSLive;"
	end if


	if isObject(session("acs_conn")) then
		set acs_conn=session("acs_conn")
	else
		set acs_conn=server.createobject("ADODB.Connection")
		session("acs_conn")=acs_conn
		'acs_conn.open "Provider=SQLOLEDB.1; Data Source=10.0.1.30; Initial Catalog=AltusInSiteLive; User ID=altusinsitelive; Password=H3Ax.4em;"
		acs_conn.open "Provider=SQLOLEDB.1; Data Source=sql-qa-use1.altusinsite.com; Initial Catalog=AltusInSiteLive; User ID=ITSLive; Password=ITSLive;"
		'acs_conn.open "Provider=SQLOLEDB.1; Data Source=192.168.0.2; Initial Catalog=AltusInSiteLive; User ID=ITSLive; Password=ITSLive;"
	end if


	g_ITS_REPORT_CODE=5859
	g_FIRST_PERIOD_ID=72
	g_MIN_YEAR=2001 ' used to be 2000
	g_MIN_PERIOD_ID=76 ' used to be 72
	g_MAX_PERIOD_ID=175
	g_CURRENT_PERIOD_ID=175
	g_HISTORY_PERIOD_ID_ARR="76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,121,122,123,124,125,126,127,128,129,130,131,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,161,162,163,164,165,166,167,168,169,170,171,172,173,174,175"
	g_ITS_SUPER_USER="gmont"

	g_ITS_GUEST_UID="itsguest"
	g_ITS_GUEST_PWD="itsguest"
%>
