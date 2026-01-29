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
require_once('templates/class.cForm_survey_1.php');

require_once('classes/class.User.php');
require_once(_FORMAPI_PATH . 'classes/class.cDB.mssql.php');

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

// check for user logged in
$user = new User($db_array);
if (!$user->IsLoggedIn())
{
	header("Location: index.php");
}

// check if a questionnaire has been completed
if (!$user->CheckCompletedQuestionnaire())
{
	header("Location: main.php?err1=true");
	die();
}

// check if survey has already been closed
if ($user->GetSurveyCompletionDate() != NULL)
{
	header("Location: main.php?err2=true");
	die();
}

// close survey
$user->CloseSurvey();

$server_date = date("n/j/Y g:i:s A");



// set e-mail data
//$subject = _MAIL_SUBJECT_SUBMITTED;
//$message_its = "A survey has been completed by " . $user->GetName() . " on $server_date.";
//$message_user = _FORM_SUBMITTED_THANK_YOU . '.';

//$driver_export_email_its = new cDriverExportEmail(_EMAIL_FROM, _EMAIL_TO, $subject);
//$driver_export_email_its->FlagHideLiterals = true;
//$driver_export_email_its->SetHeader($message_its);

//$driver_export_email_user = new cDriverExportEmail(_EMAIL_FROM, $user->GetEmail(), $subject);
//$driver_export_email_user->FlagHideLiterals = true;
//$driver_export_email_user->SetHeader($message_user);

// build and execute loop for e-mailing
//$previous_array = $user->GetAnsweredQuestionnaires();

//foreach ($previous_array as $key => $value)
//{
//	// WARNING - qid IS A GLOBAL VARIABLE!!!
//	$qid = $key;
//
//	if ($value == 'yes')
//	{
//		//print "<br><br><br><hr><hr><hr>";
//		//print "key=" . $key . '<br>';
//
//		// initialize form
//		$converterqid = new ConverterQid($qid);
//		$form_title = $converterqid->GetFormTitle();
//
//		$form = new cForm_survey_1('form_1', ucwords($form_title));
//
//		$response_id = $user->GetResponseID($qid);
//
//		$driver_import_db = new cDriverImportSql_Model1($db_array, 'Response', 'ID', $response_id);
//
//		$form->Import($driver_import_db);
//		$form->Export($driver_export_email_its);
//		$form->Export($driver_export_email_user);
//
//		//print "qid=" . $qid . "<br>";
//		//print "form_title=" . $form_title . "<br>";
//		//print "form->GetTitle()=" . $form->GetTitle() . "<br>";
//		//print "response_id=" . $response_id . "<br>";
//
//	}
//}

DisplayHeader();
?>
<?php echo _FORM_SUBMITTED; ?>
<br />
<div align="center">
<form action="main.php" method="get">
<input type="submit" value="<?php echo _MAIN_MENU; ?>" />
</form>
<br /><br />
<form action="logout.php" method="get">
<input type="submit" value="<?php echo _LOGOUT; ?>" />
</form>
</div>

<?php
DisplayFooter();


function DisplayHeader()
{
	require_once('include/page_open.php');
}

function DisplayFooter()
{
	require_once('include/page_close.php');
}
?>