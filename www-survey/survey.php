<?php
// DK - Oct 22, 2008

session_start();

if($_SERVER['PHP_SELF'] == "/survey.php/survey.php")
{
	$_SERVER['PHP_SELF']="/survey.php";
}

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

require_once('templates/class.cForm_survey_1.php');


$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

// WARNING - qid IS A GLOBAL VARIABLE!!!
$qid = 1;

$user = new User($db_array);
if (!$user->IsLoggedIn())
{
	header("Location: index.php");
}



// define contexts
$context_web_active = new cContextWebformStandard(_DIR_TEMPLATE_SURVEY_ACTIVE);
$context_web_hidden = new cContextWebformStandard(_DIR_TEMPLATE_SURVEY_HIDDEN);
$context_web_review = new cContextWebformStandard(_DIR_TEMPLATE_SURVEY_REVIEW);


// initialize form
$form_title = "Form Title";
$form_message = "Provide your responses in the boxes below then click the <strong>Next Question</strong> button. This will also save your responses.";

$form = new cForm_survey_1('form_1', ucwords($form_title));

//DisplayAllFields($form);


if (isset($_GET['preview']))
{
	$form->SetContext($context_web_active);
	DisplayHeader();
	print("<h2>THIS VERSION OF THE SURVEY IS FOR PREVIEW OR PRINTING ONLY.</h2>");
	print("<p>To complete the online survey, follow the <strong>Main Menu</strong> link from the navigation banner at the top of this page and then follow the <strong>ITS Survey for " . _QTR_ID . " " . _QTR_YR . "</strong> link to begin.</p><p>If you would prefer to complete a 'paper' version of the survey, please print this page and submit to your Altus Group surveyor.</p><p>If you need any assistance, please follow the <strong>Help With Survey</strong> link from the navigation banner at the top of this page.<br/>&nbsp;</p>");
	$form->Display();
	print("<p align=\"center\"><a href=\"main.php\">Return to the main menu</a></p>");
	DisplayFooter();
	die();
}



// initalize drivers
$driver_import_webrequest = new cDriverImportWebRequest($_REQUEST);

// check for existing response and set drivers accordingly
$response_id = $user->GetResponseID($qid);
// $response_id of -1 means new record

//$driver_export_db = new cDriverExportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2008_Q4', 'ID', $response_id);
//$driver_import_db = new cDriverImportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2008_Q4', 'ID', $response_id);

//$driver_export_db = new cDriverExportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2009_Q1', 'ID', $response_id);
//$driver_import_db = new cDriverImportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2009_Q1', 'ID', $response_id);

//$driver_export_db = new cDriverExportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2009_Q2', 'ID', $response_id);
//$driver_import_db = new cDriverImportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2009_Q2', 'ID', $response_id);

//$driver_export_db = new cDriverExportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2009_Q3', 'ID', $response_id);
//$driver_import_db = new cDriverImportSql_Model1($db_array, 'dbo.tmp_ResponseFlat_2009_Q3', 'ID', $response_id);

$driver_export_db = new cDriverExportSql_Model1($db_array, 'tmp_ResponseFlat_' . _QTR_YR . '_' . _QTR_ID, 'ID', $response_id);
$driver_import_db = new cDriverImportSql_Model1($db_array, 'tmp_ResponseFlat_' . _QTR_YR . '_' . _QTR_ID, 'ID', $response_id);


$driver_export_email = new cDriverExportEmail(_EMAIL_FROM, $user->GetEmail(), _MAIL_SUBJECT_SAVED);
$driver_export_email->FlagHideLiterals = true;
$driver_export_email->SetHeader(_MAIL_INTRO);


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
	for ($i = 2; $i <= $formwizard->GetStepCount(); $i++)
	{
		if (isset($_REQUEST[$i])) { $action = $i - 1; break; }
	}
}


if (isset($_GET['review']))
{
	if ($_GET['review'] == '1')
	{
		$action = 'review';
	}
}

$survey_completion_date = $user->GetSurveyCompletionDate();
if (($survey_completion_date != NULL) && ($action != 'save'))
{
	$action = 'review';
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
}


// submit to the database
if ($action == 'save')
{
	header("Location: main.php");
	die();
}


// Save current data to the database

if (($_POST) && (!isset($_GET['review'])) && ($survey_completion_date == NULL) && ($validation_error == ''))
{
	$form->FindSheet('sheet_1')->AddInput(new cInputHidden('IP_addr', $_SERVER['REMOTE_ADDR']));
	$form->FindSheet('sheet_1')->AddInput(new cInputHidden('OrgID', $user->GetUserID()));
	$form->FindSheet('sheet_1')->AddInput(new cInputHidden('VersionID', $qid));
	$form->FindSheet('sheet_1')->AddInput(new cInputHidden('DateModified', date("n/j/Y g:i:s A")));

	$formwizard->Export($driver_export_db);

	$form->FindSheet('sheet_1')->RemoveInput('IP_addr');
	$form->FindSheet('sheet_1')->RemoveInput('OrgID');
	$form->FindSheet('sheet_1')->RemoveInput('VersionID');
	$form->FindSheet('sheet_1')->RemoveInput('DateModified');

	if (!($form->IsCompleted())) $user->DeleteQuestionnaire($qid); // remove empty questionnaire

	// Commented out by client request - CS 8/18/2008
	// $formwizard->Export($driver_export_email);
}


DisplayHeader();

if ($action == 'review') // display the review page
{
	if ($survey_completion_date == NULL)
	{
		$formwizard->SetActiveStep(_FW_REVIEW_ALLSTEPS_CHANGE);
		$message = _PLEASE_REVIEW;
	}
	else
	{
		$formwizard->SetActiveStep(_FW_REVIEW_ALLSTEPS_NOCHANGE);
		$message = _ALREADY_SUBMITTED;
	}

	print '<p>' . $message . '</p>';

	$formwizard->Display();

	print '<p align="center">To submit your completed survey, please click the <strong>Main Menu</strong> button and follow the <strong>SUBMIT YOUR RESPONSES</strong> link.</p>';
}
else // display one of the steps
{
	$formwizard->SetActiveStep($action);
	$formwizard->SetReviewState(true);

	$step = $action + 1;
	$message = $form_message;

	print '<p>' . $message . '</p>';
	print '<p>' . _CURRENTLY_ON . ' ' . $step . ' ' . _OF . ' ' . $formwizard->GetStepCount() . '&nbsp;&nbsp;';
	DisplayProgressBar($step, $formwizard->GetStepCount());
	print '</p>';

	    if (_QTR_ID == "Q1") DisplaySection_Q1($form->GetSheet($step-1));
	elseif (_QTR_ID == "Q2") DisplaySection_Q2($form->GetSheet($step-1));
	elseif (_QTR_ID == "Q3") DisplaySection_Q3($form->GetSheet($step-1));
	elseif (_QTR_ID == "Q4") DisplaySection_Q4($form->GetSheet($step-1));

	print $validation_error;

	$formwizard->Display();
}

DisplayFooter();



function DisplayAllFields($form)
{
	$i = 0;
	while ($sheet = $form->GetSheet($i))
	{
		$j = 0;
		while ($input = $sheet->GetInput($j))
		{
			echo $input->GetName() . '<br/>';
			$j++;
		}
		$i++;
	}
}


function DisplayProgressBar($step, $total)
{

	for ($i = 1; $i <= $total; $i++)
	{
		$colour = 'yellow';
		if ($i <= $step)
		{
			$colour = 'purple';
		}
		print '<img src="images/block_' . $colour . '.gif" width="10" height="10" alt="" />';
	}
}


function DisplaySection_Q1($sheet)
{
	$sheet_name=$sheet->GetName();

	//echo $sheet_name;

	switch ($sheet_name)
	{
		case 'sheet_1':
			print '<h2>I) Altus InSite Product / Market Barometer</h2>';
			break;

		case 'sheet_2':
		case 'sheet_3':
		case 'sheet_4':
		case 'sheet_5':
		case 'sheet_6':
		case 'sheet_7':
		case 'sheet_8':
		case 'sheet_9':
		case 'sheet_10':
		case 'sheet_11':
		case 'sheet_12':
			print '<h2>II) Valuation Parameters – Office</h2>';
			break;

		case 'sheet_13':
		case 'sheet_14':
		case 'sheet_15':
		case 'sheet_16':
		case 'sheet_17':
		case 'sheet_18':
		case 'sheet_19':
			print '<h2>III) Valuation Parameters – Retail</h2>';
			break;

		case 'sheet_20':
		case 'sheet_21':
		case 'sheet_22':
		case 'sheet_23':
			print '<h2>IV) Valuation Parameters – Industrial</h2>';
			break;

		case 'sheet_24':
		case 'sheet_25':
		case 'sheet_26':
		case 'sheet_27':
			print '<h2>V) Valuation Parameters – Multi-Unit Residential</h2>';
			break;

		case 'sheet_28':
		case 'sheet_29':
			print '<h2>VI) Capital Sources – Equity Market Focus</h2>';
			break;

		case 'sheet_30':
		case 'sheet_31':
			print '<h2>VII) Miscellaneous</h2>';
			break;
	}
}

function DisplaySection_Q2($sheet)
{
	$sheet_name=$sheet->GetName();

	//echo $sheet_name;

	switch ($sheet_name)
	{
		case 'sheet_1':
			print '<h2>I) Altus InSite Product / Market Barometer</h2>';
			break;

		case 'sheet_2':
		case 'sheet_3':
		case 'sheet_4':
		case 'sheet_5':
		case 'sheet_6':
		case 'sheet_7':
		case 'sheet_8':
		case 'sheet_9':
		case 'sheet_10':
		case 'sheet_11':
		case 'sheet_12':
		case 'sheet_13':
			print '<h2>II) Valuation Parameters – Office</h2>';
			break;

		case 'sheet_14':
		case 'sheet_15':
		case 'sheet_16':
		case 'sheet_17':
		case 'sheet_18':
			print '<h2>III) Valuation Parameters – Retail</h2>';
			break;

		case 'sheet_19':
		case 'sheet_20':
		case 'sheet_21':
		case 'sheet_22':
			print '<h2>IV) Valuation Parameters – Industrial</h2>';
			break;

		case 'sheet_23':
		case 'sheet_24':
			print '<h2>V) Valuation Parameters – Multi-Unit Residential</h2>';
			break;

		case 'sheet_25':
		case 'sheet_26':
		case 'sheet_27':
			print '<h2>VI) Capital Sources – Equity Market Focus</h2>';
			break;

		case 'sheet_28':
		case 'sheet_29':
		case 'sheet_30':
		case 'sheet_31':
		case 'sheet_32':
		case 'sheet_33':
		case 'sheet_34':
			print '<h2>VII) Miscellaneous</h2>';
			break;
	}
}

function DisplaySection_Q3($sheet)
{
	$sheet_name=$sheet->GetName();

	//echo $sheet_name;

	switch ($sheet_name)
	{
		case 'sheet_1':
			print '<h2>I) Altus InSite Product / Market Barometer</h2>';
			break;

		case 'sheet_2':
		case 'sheet_3':
		case 'sheet_4':
		case 'sheet_5':
		case 'sheet_6':
		case 'sheet_7':
		case 'sheet_8':
		case 'sheet_9':
		case 'sheet_10':
		case 'sheet_11':
		case 'sheet_12':
		case 'sheet_13':
			print '<h2>II) Valuation Parameters – Office</h2>';
			break;

		case 'sheet_14':
		case 'sheet_15':
		case 'sheet_16':
		case 'sheet_17':
		case 'sheet_18':
		case 'sheet_19':
			print '<h2>III) Valuation Parameters – Retail</h2>';
			break;

		case 'sheet_20':
		case 'sheet_21':
		case 'sheet_22':
		case 'sheet_23':
			print '<h2>IV) Valuation Parameters – Industrial</h2>';
			break;

		case 'sheet_24':
		case 'sheet_25':
		case 'sheet_26':
		case 'sheet_27':
			print '<h2>V) Valuation Parameters – Multi-Unit Residential</h2>';
			break;

		case 'sheet_28':
		case 'sheet_29':
		case 'sheet_30':
			print '<h2>VI) Capital Sources – Debt Market Focus</h2>';
			break;

		case 'sheet_31':
		case 'sheet_32':
			print '<h2>VII) Miscellaneous</h2>';
			break;
	}
}

function DisplaySection_Q4($sheet)
{
	$sheet_name=$sheet->GetName();

	//echo $sheet_name;

	switch ($sheet_name)
	{
		case 'sheet_1':
			print '<h2>I) Altus InSite Product / Market Barometer</h2>';
			break;

		case 'sheet_2':
		case 'sheet_3':
		case 'sheet_4':
		case 'sheet_5':
		case 'sheet_6':
		case 'sheet_7':
		case 'sheet_8':
		case 'sheet_9':
		case 'sheet_10':
		case 'sheet_11':
		case 'sheet_12':
			print '<h2>II) Valuation Parameters – Office</h2>';
			break;

		case 'sheet_13':
		case 'sheet_14':
		case 'sheet_15':
		case 'sheet_16':
		case 'sheet_17':
		case 'sheet_18':
		case 'sheet_19':
			print '<h2>III) Valuation Parameters – Retail</h2>';
			break;

		case 'sheet_20':
		case 'sheet_21':
		case 'sheet_22':
		case 'sheet_23':
			print '<h2>IV) Valuation Parameters – Industrial</h2>';
			break;

		case 'sheet_24':
		case 'sheet_25':
		case 'sheet_26':
		case 'sheet_27':
			print '<h2>V) Valuation Parameters – Multi-Unit Residential</h2>';
			break;

		case 'sheet_28':
		case 'sheet_29':
		case 'sheet_30':
			print '<h2>VI) Capital Sources – Equity Market Focus</h2>';
			break;

		case 'sheet_31':
			print '<h2>VII) Miscellaneous</h2>';
			break;
	}
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