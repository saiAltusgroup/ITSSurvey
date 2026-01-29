<?php
// DK - Oct 22, 2008

session_start();
session_destroy();
header("Location: index.php");
?>