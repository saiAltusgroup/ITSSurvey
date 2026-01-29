<?php
	$qid=1;
	if (isset($_GET['qid'])) $qid=$_GET['qid'];
?>

<form action="<?php echo $_SERVER['PHP_SELF'] . '?qid=' . $qid; ?>" method="post" class="fa_active_form" name="mainform">
