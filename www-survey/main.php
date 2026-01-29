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

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

$user = new User($db_array);
if (!$user->IsLoggedIn())
{
	header("Location: index.php");
}
//if (!$user->GetUserSaved())
//{
//	header("Location: user.php");
//}

require_once('include/page_open.php');

$survey_completion_date = $user->GetSurveyCompletionDate();
?>

<h2 align="center"><?php echo _MAIN_MENU; ?></h2>

<?php
if (isset($_GET['err1']))
{
	if ($_GET['err1'] == 'true') { print '<p class="error"><strong>' . _ERROR_MUST_COMPLETE_FORM . '</strong></p>'; }
}
if (isset($_GET['err2']))
{
	if ($_GET['err2'] == 'true') { print '<p class="error"><strong>' . _ERROR_ALREADY_COMPLETE . '</strong></p>'; }
}

$login_count = $user->GetLoginCount();
$login_count_plural = '';
if ($login_count != 1)
{
	$login_count_plural = 's';
}
?>

<fieldset>
<legend><?php echo _WELCOME; ?> <?php echo htmlentities($user->GetName());?></legend>
<div class="descr">
<?php
if (strlen($user->GetPreviousLoginDate()) > 0)
{
?>
<?php echo _LAST_LOGGED_IN; ?> <strong><?php echo $user->GetPreviousLoginDate(); ?></strong>.<br />
<?php
}
?>
<?php echo _LOGIN_COUNT; ?> <strong><?php echo $login_count; ?></strong>
<br /><br />Your user profile has been successfully created and you are now ready to begin completing the Investment Trends Survey for <?php echo _QTR_ID; ?> <?php echo _QTR_YR; ?>.
</div>
<ul>
<li><a href="user.php"><?php echo _UPDATE_PROFILE; ?></a></li>
<!--<li><a href="logout.php"><?php echo _LOGOUT; ?></a></li>-->
</ul>

<!--
<?php echo 'Expertise: ' . $user->GetEXPERTISE() . '<br />'; ?>
<?php echo 'Markets: ' . $user->GetMARKETS() . '<br />'; ?>
-->

</fieldset>

<?php
$previous_array = $user->GetAnsweredQuestionnaires();
$pencil_array = array();

foreach ($previous_array as $key => $value)
{
	if ($value == 'yes')
	{
		$pencil_array[$key][0] = '<img src="images/pencil.gif" alt="pencil" width="15" height="24" />';
		$pencil_array[$key][1] = '&amp;review=1';
	}
	else
	{
		$pencil_array[$key][0] = '<img src="images/blank.gif" alt="" width="15" height="24" />';
		$pencil_array[$key][1] = '';
	}
}
?>
<fieldset>
<legend><?php echo _SURVEY; ?></legend>
<div class="descr">
<?php
if ($survey_completion_date == NULL)
{
	print _SURVEY_COMPLETION_NULL;
}
else
{
	print _SURVEY_COMPLETION_1 . $survey_completion_date . _SURVEY_COMPLETION_2;
	print _SURVEY_COMPLETION_3;
}
?>
</div>

<ul style="margin-top: 0;">
<li><a href="survey.php?qid=1<?php echo $pencil_array[1][1]; ?>"><?php echo _SURVEY_CURRENT; ?></a> <?php echo $pencil_array[1][0]; ?>
<br /><small><br />To preview the survey before filling it out, <a href="survey.php?preview=1&amp;qid=1">click here for a printable version of the <?php echo _SURVEY_CURRENT; ?></a>.</small>
</li>
	<br />
	<li>
		<?php require_once('templates/index.' . $lang . '.2.1.php'); ?>
	</li>
</ul>
</fieldset>
<br />
<?php
if ($survey_completion_date == NULL)
{
	$link_1 = '';
	$link_2 = '';
	if ($user->CheckCompletedQuestionnaire())
	{
		$link_1 = '<a href="final_check.php">';
		$link_2 = '</a>';
	}

	print '<div align="center">' . $link_1 . '<img src="images/submit.gif" width="90" height="72" alt="Final Submit" border="0" />' . $link_2 . '<br /><strong>' . $link_1 . _FINAL_SUBMIT . $link_2 . '</strong><br />' . _SUBMIT_ONLY_WHEN . '</div>';
}

require_once('include/page_close.php');
?>
