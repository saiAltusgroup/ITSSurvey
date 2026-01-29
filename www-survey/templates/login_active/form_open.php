<form action="<?php
if($_SERVER['PHP_SELF'] == "/index.php/index.php")
{
    $_SERVER['PHP_SELF']="/index.php";
}
    echo $_SERVER['PHP_SELF']; ?>" method="post">
