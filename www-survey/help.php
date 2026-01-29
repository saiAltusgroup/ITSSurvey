<?php
// DK - Feb 3, 2014

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

require_once('include/page_open.php');

?>

<h2>Survey Assistance</h2>

<p>
If you would like assistance with completing your survey, please contact your <em>Altus Group Surveyor</em>, which you can access from your <strong>Contributor Profile Page</strong>. To access this page, you must <a href="index.php">log in</a> and click the <strong>Update Your Profile</strong> link.
</p>

<p>
	If you are experiencing any difficulty with the survey or have any questions or feedback, please contact <a href="mailto:florence.mjama@altusgroup.com?subject=Help%20with%20ITS">Florence Mjama</a> at <strong>(604) 683-1680 ext 1680</strong>.
</p>

<?php
/*

<p>
To complete the survey manually, you can download a <strong>PDF version of the survey</strong> using the links below:
</p>

<ul>
<li><a href="http://www.altusinsite.com/surveys/its/IT216-Survey-Complete.pdf">English Version</a></li>
<li><a href="http://www.altusinsite.com/surveys/its/IT216-Survey-Complete_FR.pdf">French Version</a></li>
</ul>
*/
?>
<p lang="fr-ca">
Pour assistance en fran&ccedil;ais, communiquer avec <a href="mailto:diana.pricop@groupealtus.com?subject=Investment%20Trends%20Survey">Diana Pricop</a> au <strong>(438)&nbsp;469-5484</strong>.
</p>


<?php
require_once('include/page_close.php');
?>