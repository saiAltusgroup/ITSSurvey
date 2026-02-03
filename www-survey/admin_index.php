<?php
// DK - Nov 19, 2008
// BD - Mar 10, 2009

session_start();

if (!isset($_SESSION['admin_loggedin'])) { header("Location: admin_login.php"); exit(); }

require_once('sys_include/globals.php');
require_once('templates/lang.user.en.php');

require_once('classes/class.User.php');
require_once(_FORMAPI_PATH . 'classes/class.cDB.mssql.php');
require_once('classes/class.Userlist.php');

$viewid = 1; if (isset($_GET['viewid'])) { $viewid = $_GET['viewid']; }

$sortby = 'ID';
if (isset($_GET['sortby']))
{
	$sortby = $_GET['sortby'];
	if (!in_array($sortby, array('CONTRIB', 'CONTRIB DESC', 'SURVEYOR', 'SURVEYOR DESC', 'LastLoginDate DESC', 'LastLoginDate', 'CompletedDate DESC' , 'CompletedDate')))
	{
		$sortby = 'ID';
	}
}

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

DisplayHeader();
DisplayTable($db_array, $viewid, $sortby);
DisplayFooter();

function DisplayTable($db_array, $viewid, $sortby)
{
	$userlist = new Userlist($db_array);
	$userlist->SetFilter($viewid);
	$userlist->BuildArray($sortby);

	print "<ul>";
	print ($viewid == 0 ? "<li><strong>" : "<li>") . "<a href=\"admin_index.php?viewid=0\">Show all contributors</a>" . ($viewid == 0 ? "</strong></li>" : "</li>");
	print ($viewid == 1 ? "<li><strong>" : "<li>") . "<a href=\"admin_index.php?viewid=1\">Show only contributors who have logged in during the current survey period</a>" . ($viewid == 1 ? "</strong></li>" : "</li>");
	print ($viewid == 2 ? "<li><strong>" : "<li>") . "<a href=\"admin_index.php?viewid=2\">Show only contributors who have completed the current survey</a>" . ($viewid == 2 ? "</strong></li>" : "</li>");;
	print "</ul>";

	$sort_link = 'admin_index.php?viewid=' . $viewid . '&amp;sortby=';

	$sort_link_contrib =  $sort_link . (($sortby=='CONTRIB') ? 'CONTRIB+DESC' : 'CONTRIB');
	$sort_link_surveyor = $sort_link . (($sortby=='SURVEYOR') ? 'SURVEYOR+DESC' : 'SURVEYOR');
	$sort_link_last_login = $sort_link . (($sortby=='LastLoginDate DESC') ? 'LastLoginDate' : 'LastLoginDate+DESC');
	$sort_link_completed = $sort_link . (($sortby=='CompletedDate DESC') ? 'CompletedDate' : 'CompletedDate+DESC');

	print <<< EOF
<table border="0" cellpadding="5" cellspacing="1" width="100%" id="admin_table">
	<tr align="center">
		<th>ID</th>
		<th align="left"><a href='$sort_link_contrib'>Contributor</a></th>
		<th>Username</th>
		<th>Password</th>
		<th align="left">Phone</th>
		<th align="left">Email</th>
		<th align="left"><a href='$sort_link_surveyor'>Surveyor</a></th>
		<th><a href='$sort_link_last_login'>Last Login Date</a></th>
		<th><a href='$sort_link_completed'>Survey Completion Date</a></th>
		<th>Any Answers?</th>
	</tr>
EOF;

	$user = new User($db_array);

	foreach ($userlist->GetArray() as $user_id)
	{
//		echo $user_id . "<BR>";

		$user->RefreshById($user_id);

		$userid = htmlentities($user->GetUserID());
		$name = htmlentities($user->GetName());
		$phone = htmlentities($user->GetPhone());
		$email = htmlentities($user->GetEmail());
		$surveyor = htmlentities($user->GetSurveyor());
		$username = htmlentities($user->GetUsername());
		$password = htmlentities($user->GetPassword());
		$login_last = htmlentities($user->GetLastLoginDate());
		$date = htmlentities($user->GetSurveyCompletionDate());

		$answered_array = $user->GetAnsweredQuestionnaires();
		$view_array = array();

		foreach ($answered_array as $key => $value)
		{
			if ($value == 'yes')
			{
				$view_array[$key] = '<strong>Yes</strong> <!--<a href="admin_survey.php?qid=' . $key . '&amp;userid=' . $userid . '">View</a>-->';
			}
			else
			{
				$view_array[$key] = 'No';
			}
		}

		print <<< EOF
	<tr align="center" valign="top">
		<td>$userid</td>
		<td align="left"><a href="user.php?userid_f234=$userid">$name</a></td>
		<td align="left">$username</td>
		<td align="left">$password</td>
		<td align="left">$phone</td>
		<td align="left">$email</td>
		<td align="left">$surveyor</td>

		<td>$login_last</td>
		<td style="font-weight:bold;">$date</td>
		<td>{$view_array[1]}</td>
	</tr>
EOF;
	}

	print "</table>";
}

function DisplayHeader()
{
	require_once('include/page_open.php');
	print '<h2 align="center">ITS Survey Administration Login</h2>';
}

function DisplayFooter()
{
	require_once('include/page_close.php');
}

?>