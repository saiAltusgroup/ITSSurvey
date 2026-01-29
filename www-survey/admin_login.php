<?php

session_start();

require_once('sys_include/globals.php');
require_once('templates/lang.en.php');

// define contexts
$context_web_active = new cContextWebformStandard(_DIR_TEMPLATE_LOGIN_ACTIVE);

// initalize drivers
$driver_import_webrequest = new cDriverImportWebRequest($_REQUEST);

// initialize form
$form = new cForm_login_2('form_2', _FORM_LOGIN);

// validate
if ($_POST)
{
	$form->Import($driver_import_webrequest);

	if ($form->Validate())
	{
		$_SESSION['admin_loggedin'] = 'yes';
		header("Location: admin_index.php");
	}
	else
	{
		header("Location: admin_login.php");
	}
}


DisplayHeader();
$form->SetContext($context_web_active);
$form->Display();
DisplayFooter();


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
