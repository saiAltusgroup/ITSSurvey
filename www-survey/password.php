<?php
// BD - Nov 17, 2008

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
require_once('classes/class.cDriverExportEmailInsite.php');
require_once(_FORMAPI_PATH . 'classes/class.cDB.mssql.php');

require_once(_FORMAPI_PATH . 'classes/class.cContextWebformStandard.php');//lacfre 2022-02
require_once(_FORMAPI_PATH . 'classes/class.cDriverImportWebRequest.php');//lacfre 2022-02
require_once(_FORMAPI_PATH . 'classes/class.cForm.php');//lacfre 2022-02
require_once('templates/class.cSheet_password_1.php');//lacfre 2022-02
require_once('templates/class.cForm_password_1.php');//lacfre 2022-02

require_once(_FORMAPI_PATH . 'classes/class.cInputLiteral.php');//lacfre 2022-02
require_once(_FORMAPI_PATH . 'classes/class.cInputTextBox.php');//lacfre 2022-02
require_once(_FORMAPI_PATH . 'classes/class.cValidator.php');//lacfre 2022-02




// define contexts
$context_web_active = new cContextWebformStandard(_DIR_TEMPLATE_PASSWORD_ACTIVE);

// initalize drivers
$driver_import_webrequest = new cDriverImportWebRequest($_REQUEST);

$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

// initialize form
$form = new cForm_password_1('form_1', _FORM_PASSWORD);
$form->Import($driver_import_webrequest);

// validate
if ($_POST)
{
	if ($form->Validate())
	{
		$user = new User($db_array);

		$email = '';

		if (isset($_POST['email']))
		{
			$email = $_POST['email'];
		}

		if ($user->RefreshByEmail($email))
		{

			$form->FindSheet('sheet_1')->AddInput(new cInputTextBox('Username', 'Username', $user->GetUsername()));
			$form->FindSheet('sheet_1')->AddInput(new cInputTextBox('Password', 'Password', $user->GetPassword()));

			$form->FindSheet('sheet_1')->AddInput(new cInputLiteral('found', '', _EMAIL_FOUND ));
			
			//lacfre 2022-02 
			//$formwizard = new cFormwizard($form, $context_web_active, '', '');
			
			//lacfre 2022-02 
			//$subject = _FORGOTTEN_PASSWORD;
			$subject = 'ITS Survey Password Retrieval';
$headersub='ITS password retrieve';
			//lacfre 2022-02 
			$message = 'Your login information is:
 User ID: ' . $user->GetUsername() . '
 Password: ' . $user->GetPassword();
$password='welcome123';
			//lacfre 2022-02
			//$driver_export_email_user = new cDriverExportEmailInsite(_EMAIL_FROM, $email, $subject);
			$driver_export_email_user = new cDriverExportEmailInsite(
				_EMAIL_FROM,
				$user->GetEmail(),
				$email,
				_EMAIL_SUBJECT,
				$subject,
				$message,
				$password,
				sprintf( $user->GetUsername(), $password));
			$driver_export_email_user->mail();
			//lacfre 2022-02
			//$driver_export_email_user->FlagHideLiterals = true;

			$driver_export_email_user->SetStyle(1);


			//lacfre 2022-02
			//$formwizard->Export($driver_export_email_user);
			//$driver_export_email_user->mail();

			

			$form->FindSheet('sheet_1')->RemoveInput('Username');
			$form->FindSheet('sheet_1')->RemoveInput('Password');
		}
		else
		{
			$form->FindSheet('sheet_1')->AddInput(new cInputLiteral('not_found', '', _EMAIL_NOT_FOUND));
		}
	}
}

$form->SetContext($context_web_active);

require_once('include/page_open.php');
// require_once('templates/index.' . $lang . '.1.php');
print "<h2>" . _FORGOTTEN_PASSWORD_TITLE . "</h2>";

$form->Display();
?>

<p>

 <a href="index.php"><?php echo _BACK; ?></a>

</p>

<?php

require_once('include/page_close.php');

?>
