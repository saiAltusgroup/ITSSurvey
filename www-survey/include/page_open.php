<?php

print <<< EOF
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
    "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="en-ca">
<head>
	<title>ITS Survey</title>
	<link rel="stylesheet" href="include/screen.css" type="text/css" />
	<script type="text/javascript" src="include/main.js"></script>
</head>
<body>

<div class="mainlayout">

<div class="stationary">
	<div class="navheader">
		<a href="help.php">Help With Survey</a> |
		<a href="user.php">User Profile</a> |
		<a href="mailto:support@altusinsite.com?Subject=Help%20with%20Investment%20Trends%20Survey">Support</a> |
		<a href="main.php">Main Menu</a> |
		<a href="logout.php">Logout</a>
	</div>
</div>

<div class="content">

EOF;
?>

<h1><?php echo _QTR_ID; ?> <?php echo _QTR_YR; ?> Survey</h1>
