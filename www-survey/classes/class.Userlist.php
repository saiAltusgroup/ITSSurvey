<?php
// DK - Nov 19, 2008
// BD - Mar 10, 2009, adding sorting
// DK - May 22, 2013, added additiomal filtering (the SetFilter method), re-implemented the legacy SetCompletionFilter method using the new SetFilter method

class Userlist
{
	private $DB;
	private $UserIDArray;
	private $Filter;

	public function __construct($db_array)
	{
		$this->DB = new cDB;
		$this->DB->connect($db_array['server'], $db_array['username'], $db_array['password']);
		$this->DB->select_db($db_array['database']);
		$this->Filter = 2;
		$this->SortBy = 'ID';
		$this->UserIDArray = array();
	}

	public function __destruct()
	{
		$this->DB->close();
	}

	public function BuildArray($sortby)
	{
		$where_sql = '';

		if ($this->Filter == 2) /* Only contributors who have completed the current survey */
		{
			$where_sql = 'WHERE CompletedDate IS NOT NULL';
		}
		else if ($this->Filter == 1) /* Only contributors who have logged in during the current survey period */
		{
			/*
			Q1 - February 1 - March 31
			Q2 - May 1 - June 30
			Q3 - August 1 - September 30
			Q4 - November 1 - December 31
			*/

			     if (_QTR_ID == 'Q1') { $from =  '2/1/'; $to =  '3/31/'; }
			else if (_QTR_ID == 'Q2') { $from =  '5/1/'; $to =  '6/30/'; }
			else if (_QTR_ID == 'Q3') { $from =  '8/1/'; $to =  '9/30/'; }
			else                      { $from = '11/1/'; $to = '12/31/'; }

			$from .= _QTR_YR;
			$to .= _QTR_YR;
			$where_sql = 'WHERE LastLoginDate >=\'' . $from . '\' AND LastLoginDate <= \'' . $to . '\''; 
		}

		$this->SortBy = $sortby;

		$sql = "SELECT ID FROM dat_Contributor $where_sql ORDER BY " . $this->SortBy;

		$result = $this->DB->query($sql);

		while ($line = $this->DB->fetch_array($result))
		{
			$this->UserIDArray[] = $line['ID'];
		}
	}

	public function GetArray()
	{
		return $this->UserIDArray;
	}

	public function SetFilter($filter)
	{
		$this->Filter = $filter;
	}

	public function SetCompletionFilter($filter) /* old */
	{
		if ($filter) $this->Filter = 2; else $this->Filter = 0;
	}
}
?>