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

require_once('classes/class.User.php');
require_once(_FORMAPI_PATH . 'classes/class.cDB.mssql.php');

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

$user = new User($db_array);
if (!$user->IsLoggedIn())
{
	header("Location: index.php");
}

if (!$user->CheckCompletedQuestionnaire())
{
	header("Location: main.php?err1=true");
}

require_once('include/page_open.php');
?>

<h2 align="center"><?php echo _FINAL_SUBMISSION_CHECK; ?></h2>

<?php echo _FINAL_SUBMISSION_CHECK_TEXT; ?>

<div align="center">
<form action="final_submit.php" method="get">
<input type="submit" value="<?php echo _SUBMIT_NOW; ?>" />
</form>
</div>

<?php
require_once('include/page_close.php');
?>