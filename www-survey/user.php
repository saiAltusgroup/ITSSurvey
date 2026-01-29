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
require_once('templates/lang.user.' . $lang . '.php');

require_once('classes/class.User.php');
require_once(_FORMAPI_PATH . 'classes/class.cDB.mssql.php');

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

$user = new User($db_array);

$mode='edit';
if (!($user->IsLoggedIn())) $mode='insert';

//print $user->GetUserID();


// require files
require_once('templates/class.cForm_user_1.php');

// define contexts
$context_web_active = new cContextWebformStandard(_DIR_TEMPLATE_USER_ACTIVE);
$context_web_hidden = new cContextWebformStandard(_DIR_TEMPLATE_SURVEY_HIDDEN);
$context_web_review = new cContextWebformStandard(_DIR_TEMPLATE_SURVEY_REVIEW);


// initalize drivers

$driver_import_webrequest = new cDriverImportWebRequest($_REQUEST);

if (isset($_GET['userid_f234']))
{
	$userid=  $_GET['userid_f234'];
	$driver_import_db = new cDriverImportSql_Model1($db_array, 'dat_Contributor', 'ID', $userid);
}
else
{
	$driver_import_db = new cDriverImportSql_Model1($db_array, 'dat_Contributor', 'ID', $user->GetUserID());
}

// initialize form
$form = new cForm_user_1('form_1', _FORM_1);

// initialize formwizard
$formwizard = new cFormwizard($form, $context_web_active, $context_web_hidden, $context_web_review);

if ($_POST)
{
	$formwizard->Import($driver_import_webrequest);
}
else
{
	$formwizard->Import($driver_import_db);
}

// determine current step
if (isset($_REQUEST['review']))
{
	$action = 'review';
}
elseif (isset($_REQUEST['save']))
{
	$action = 'save';
}
else
{
	$action = '0';
}

if (isset($_GET['userid_f234']))
{
	$action = 'display_only';
}

$validation_error = '';
// validate the step
if ($_POST)
{
	if (is_numeric($_REQUEST['from_step']))
	{
		if (!$formwizard->ValidateStep($_REQUEST['from_step']))
		{
			$action = $_REQUEST['from_step'];
			$validation_error = '<p class="error"><strong>' . _ERROR_FORM_VALIDATION . '</strong></p>';
		}
	}

	if (isset($_POST['Username'])) $username = $_POST['Username'];
	if (isset($_POST['Password'])) $password = $_POST['Password'];
}

//print $user->GetUserID();

// display the review page
if ($action == 'review')
{
	       $sheet=$form->FindSheet('sheet_1');

	       $formwizard->SetActiveStep(_FW_REVIEW_ALLSTEPS_CHANGE);

	       $sheet->AddInput(new cInputHidden('IP_addr', $_SERVER['REMOTE_ADDR']));
	       $sheet->AddInput(new cInputHidden('DateModified', date("n/j/Y g:i:s A")));
	       $sheet->AddInput(new cInputHidden('UserSaved', 'yes'));

	       $driver_export_db = new cDriverExportSql_Model1($db_array, 'dat_Contributor', 'ID', $user->GetUserID());

	       if (!$formwizard->Export($driver_export_db))
	       {
		       print '<br /><br />';
		       require("user_fail.php");
		       die();
	       }

	       $sheet->RemoveInput('IP_addr');
	       $sheet->RemoveInput('DateModified');
	       $sheet->RemoveInput('UserSaved');

	       // Remove password field before displaying the form
	       $sheet->RemoveInput('Password');

	       if ($mode == 'edit' )
	       {
		       $user->Refresh(); // because user's email address can change
	       }
	       else
	       {
		       $user->Login($username, $password);
	       }

	       $driver_export_email = new cDriverExportEmail(_EMAIL_FROM, $user->GetEmail(), _MAIL_SUBJECT);
	       $driver_export_email->SetHeader(_MAIL_INTRO);

	       if ($user->GetUserSaved() && strlen($user->GetEmail())>0)
	       {
		       //$formwizard->Export($driver_export_email);
	       }

	       DisplayHeader();
	       print '<p>' . _PLEASE_REVIEW . '</p>';
	       $formwizard->Display();
	       DisplayFooter();
	       //print $user->GetUserID();
}
elseif ($action == 'display_only')
{
	$sheet=$form->FindSheet('sheet_1');

	$formwizard->SetActiveStep(_FW_REVIEW_ALLSTEPS_CHANGE);

	// Remove password field before displaying the form
	$sheet->RemoveInput('Password');

	DisplayHeader();
	print "<p><a href='admin_index.php'> Back </a></p>";
	$formwizard->Display();
	DisplayFooter();
}
elseif ($action == 'save')
{
	header("Location: main.php");
}
else
{
	$formwizard->SetActiveStep($action);

	if (isset($_REQUEST['review_flag']))
	{
		$formwizard->SetReviewState(true);
	}

	DisplayHeader();

	if (!$user->GetUserSaved())
	{
		print '<p>' . _USER_FIRST . '</p>';
	}
	else
	{
		print '<p>' . _USER_SAVED . '</p>';
	}

	if ( $mode == 'insert' ) print "<h2>" . _USER_HEADER_NEW . "</h2>"; else print "<h2>" . _USER_HEADER_EXISTING . "</h2>";

	print '<p>Note: All fields are required.</p>';

	print $validation_error;

	$formwizard->Display();
	
	if ($user->GetUserSaved())
	{
?>
<div align="center">
<form action="main.php" method="get">
<input type="submit" value="<?php echo _CANCEL; ?>" />
</form>
</div>
<?php
	}

	DisplayFooter();
}

function DisplayHeader()
{
	require_once('include/page_open.php');
}

function DisplayFooter()
{
	require_once('include/page_close.php');
}

?>