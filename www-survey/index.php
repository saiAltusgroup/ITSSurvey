<?php
// DK - Oct 22, 2008

session_start();

$host = $_SERVER["SCRIPT_NAME"];
$lang = 'en';

if (strpos($host, '/fr/') !== false)
{
	$lang = 'fr';
}

// require files
require_once('sys_include/globals.php');
require_once('templates/lang.' . $lang . '.php');


if (_SURVEY_CLOSED)
{
	require_once('include/page_open.php');
	require_once('templates/index.' . $lang . '.closed.php');
	require_once('include/page_close.php');
	die();
}


require_once('classes/class.User.php');
require_once(_FORMAPI_PATH . 'classes/class.cDB.mssql.php');

// define contexts
$context_web_active = new cContextWebformStandard(_DIR_TEMPLATE_LOGIN_ACTIVE);

// initalize drivers
$driver_import_webrequest = new cDriverImportWebRequest($_REQUEST);

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

// initialize form
$form = new cForm_login_1('form_1', _FORM_LOGIN);
$form->Import($driver_import_webrequest);

// validate
if ($_POST)
{
	if ($form->Validate())
	{
		$user = new User($db_array);

		$username = '';
		$password = '';

		if (isset($_POST['username']))
		{
			$username = $_POST['username'];
		}

		if (isset($_POST['password']))
		{
			$password = $_POST['password'];
		}

		if ($user->Login($username, $password))
		{
			$redirect_link = 'main.php';
		}
		else
		{
			$user->Logout();

			$redirect_link = 'index.php?err1=true';
		}

		header("Location: $redirect_link");
	}
}

$form->SetContext($context_web_active);

require_once('include/page_open.php');
require_once('templates/index.' . $lang . '.1.php');

if (isset($_GET['err1']))
{
	if ($_GET['err1'] == 'true')
	{
		print '<p class="error"><strong>' . _ERROR_LOGIN . '</strong></p>';
	}
}

$form->Display();

require_once('templates/index.' . $lang . '.2.php');
require_once('include/page_close.php');

?>