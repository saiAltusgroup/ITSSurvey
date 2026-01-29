<?php

define('_APP_PATH', 'C:\\WWW\\its\\www-survey\\');
define('_FORMAPI_PATH', 'C:\\WWW-LIB\\php\\formapi_2_0\\');

define('_DIR_TEMPLATE_SURVEY_ACTIVE', _APP_PATH . 'templates\\survey_active');
define('_DIR_TEMPLATE_SURVEY_HIDDEN', _APP_PATH . 'templates\\survey_hidden');
define('_DIR_TEMPLATE_SURVEY_REVIEW', _APP_PATH . 'templates\\survey_review');
define('_DIR_TEMPLATE_LOGIN_ACTIVE', _APP_PATH . 'templates\\login_active');
define('_DIR_TEMPLATE_USER_ACTIVE', _APP_PATH . 'templates\\user_active');
define('_DIR_TEMPLATE_PASSWORD_ACTIVE', _APP_PATH .'templates\\password_active');

//########################################################################
// Adjust below to open or close survey.


//=========== Test ==============================
define('_DB_SERVER', 'sql-qa-use1.altusinsite.com');  // Test Server
define('_DB_PASSWORD', 'ITSLive');   // Test Password
//===============================================

//=========== Production ========================
//define('_DB_SERVER', '10.0.1.30');     // Production server
//define('_DB_PASSWORD', 'dak!W4yeQ');   // Production Password
//===============================================

define('_SURVEY_CLOSED', false); // Survey is closed if set to true ... survey is open when set to false

define('_QTR_ID', 'Q3');       // Survey Quarter
define('_QTR_YR', '2025');     // Survey Year

//########################################################################



define('_DB_DATABASE', 'ITSLive');
define('_DB_USERNAME', 'ITSLive');

//lacfre 2022-02
//define('_EMAIL_FROM', 'support@altusinsite.com');
define('_EMAIL_FROM', 'info_insite@aws-ses.altusgroup.com');
define('_EMAIL_SUBJECT','ITS Survey Password Retrieval');
define('_EMAIL_TO', 'support@altusinsite.com');




function __autoload($class_name)
{
	if (is_file(_FORMAPI_PATH . 'classes/class.' . $class_name . '.php')) require_once(_FORMAPI_PATH . 'classes/class.' . $class_name . '.php');
	if (is_file(_APP_PATH . 'templates/class.' . $class_name . '.php')) require_once(_APP_PATH . 'templates/class.' . $class_name . '.php');
}

?>
