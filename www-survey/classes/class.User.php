<?php
// DK - Oct 22, 2008
// DK - Nov 19, 2008 - factorization and improvements

class User
{
	private $UserID;
	private $GUID;
	private $Name;
	private $Username;
	private $Password;
	private $Email;
	private $Surveyor;
	private $Phone;
	private $CreatedDate;
	private $LastLoginDate;
	private $LoginCount;
	private $PreviousLoginDate;
	private $SurveyCompletionDate;
	private $UserSaved;
	private $EXPERTISE;
	private $MARKETS;

	private $DB;

	public function __construct($db_array)
	{
		$this->DB = new cDB;
		$this->DB->connect($db_array['server'], $db_array['username'], $db_array['password']);
		$this->DB->select_db($db_array['database']);

		$this->Reset();
		$this->Refresh();
	}

	public function __destruct()
	{
		$this->DB->close();
	}

	protected function Reset()
	{
		$this->UserID = -1;
		$this->GUID = '';
		$this->Name = '';
		$this->Username = '';
		$this->Password = '';
		$this->Email = '';
		$this->Surveyor = '';
		$this->Phone = '';
		$this->CreatedDate = '';
		$this->LastLoginDate = '';
		$this->LoginCount = '';
		$this->PreviousLoginDate = '';
		$this->SurveyCompletionDate = '';
		$this->UserSaved = 'no';
		$this->EXPERTISE = '';
		$this->MARKETS = '';
	}

	protected function RefreshByFilter($filter)
	{
		$sql = "SELECT ID, GUID, CONTRIB, PASSWORD, SURVEYOR, PHONE, Email, Username, DateCreated, LastLoginDate, LoginCount, PreviousLoginDate, CompletedDate, UserSaved, EXPERTISE, MARKETS FROM dat_Contributor";
		$sql = $sql . $filter;

		$result = $this->DB->query($sql);
		$line = $this->DB->fetch_array($result);

		if (!$line) return false;

		$this->UserID = $line['ID'];
		$this->GUID = $line['GUID'];
		$this->Name = $line['CONTRIB'];
		$this->Password = $line['PASSWORD'];
		$this->Surveyor = $line['SURVEYOR'];
		$this->Phone = $line['PHONE'];
		$this->Email = $line['Email'];
		$this->Username = $line['Username'];
		$this->CreatedDate = $line['DateCreated'];
		$this->LastLoginDate = $line['LastLoginDate'];
		$this->LoginCount = $line['LoginCount'];
		$this->PreviousLoginDate = $line['PreviousLoginDate'];
		$this->SurveyCompletionDate = $line['CompletedDate'];
		$this->UserSaved = $line['UserSaved'];
		$this->EXPERTISE = $line['EXPERTISE'];
		$this->MARKETS = $line['MARKETS'];

		return true;
	}


	// refreshes the current user (stored in session) with up to date information - used after updating the user's record in the db
	public function Refresh()
	{
		if (!$this->IsLoggedIn()) return;

		$this->UserID = $_SESSION['user_id'];
		$this->RefreshById($this->UserID);

	}

	// refreshes the user by passing user's id, does not modify the current user (stored in session)
	public function RefreshById($userid)
	{
		$userid = $this->DB->escape_string($userid);
		$filter = " WHERE ID = '$userid'";

		return $this->RefreshByFilter($filter);
	}

	// refreshes the user by passing user's email, does not modify the current user (stored in session)
	public function RefreshByEmail($email)
	{
		$email = $this->DB->escape_string($email);
		
		//lacfre 2022-02
		$filter = " WHERE Email = '$email' ORDER BY LastLoginDate DESC";

		return $this->RefreshByFilter($filter);
	}

	// logs in the user and stores its user_id in session
	public function Login($username, $password)
	{
		if ((strlen($username) <= 0) || (strlen($password) <= 0)) return false;

		$username = $this->DB->escape_string($username);
		$password = $this->DB->escape_string($password);
		$filter = " WHERE Username = '$username' AND Password = '$password'";

		if (!$this->RefreshByFilter($filter)) return false;

		// sets the session information
		session_register('user_id');
		$_SESSION['user_id'] = $this->UserID;

		// increment login count
		$this->LoginCount++;

		// updates the user's record to show the new login
		$NewLastLoginDate = date("Y-m-d H:i:s");

		$previous_sql = '';
		if (strlen($this->LastLoginDate) > 0) $previous_sql = ", PreviousLoginDate = '" . $this->LastLoginDate . "'";

		$userid = $this->DB->escape_string($this->UserID);
		$sql = "UPDATE dat_Contributor SET LastLoginDate = '$NewLastLoginDate'" . $previous_sql . ", LoginCount = LoginCount + 1 WHERE ID = '" . $userid . "'";
		$result = $this->DB->query($sql);

		return true;
	}

	// checks to see if the user is currently logged in
	public function IsLoggedIn()
	{
		if (isset($_SESSION['user_id'])) return true;
		return false;
	}

	// logs the user out, clears session
	public function Logout()
	{
		session_destroy();
		$this->Reset();
		return true;
	}

	// returns an array indicating which questionnaires have been answered
	public function GetAnsweredQuestionnaires()
	{
		$return_array = array('', 'no' /*, 'no', 'no', 'no', 'no'*/);

		$userid = $this->DB->escape_string($this->UserID);

		// Modified by BD , Feb 27, 2009
		$tablename = 'tmp_ResponseFlat_' . _QTR_YR . '_' . _QTR_ID;

		$sql = "SELECT VersionID FROM " . $tablename . " WHERE OrgID = '" . $userid . "'";

		//$sql = "SELECT VersionID FROM tmp_ResponseFlat2 WHERE OrgID = '" . $userid . "'";

		$result = $this->DB->query($sql);

		while ($line = $this->DB->fetch_array($result))
		{
			$return_array[$line['VersionID']] = 'yes';
		}

		return $return_array;
	}

	// deletes a record from the Response table based on userid and qid
	public function DeleteQuestionnaire($qid)
	{
		$userid = $this->DB->escape_string($this->UserID);
		$qid = $this->DB->escape_string($qid);

		// Modified by BD , Feb 27, 2009
		$tablename = 'tmp_ResponseFlat_' . _QTR_YR . '_' . _QTR_ID;

		$sql = "DELETE FROM " . $tablename . " WHERE OrgID = '" . $userid . "' AND VersionID = '" . $qid . "'";
//		$sql = "DELETE FROM tmp_ResponseFlat2 WHERE OrgID = '" . $userid . "' AND VersionID = '" . $qid . "'";

		$result = $this->DB->query($sql);

		return true;
	}

	// checks if at least one questionnaire has been answered and if so, returns true
	public function CheckCompletedQuestionnaire()
	{
		$previous_array = $this->GetAnsweredQuestionnaires();
		foreach ($previous_array as $key => $value)
		{
			if ($value == 'yes') return true;
		}

		return false;
	}

	public function CloseSurvey()
	{
		$server_date = date("n/j/Y g:i:s A");
		$userid = $this->DB->escape_string($this->UserID);
		$sql = "UPDATE dat_Contributor SET CompletedDate = '$server_date' WHERE ID = '" . $userid . "'";
		$result = $this->DB->query($sql);

		return true;
	}

	// gets the ID of the record that has the appropriate Org ID and Version ID
	public function GetResponseID($qid)
	{
		$return_value = -1;

		$qid = $this->DB->escape_string($qid);
		$orgid = $this->DB->escape_string($this->UserID);


		// Modified by BD , Feb 27, 2009
		$tablename = 'tmp_ResponseFlat_' . _QTR_YR . '_' . _QTR_ID;

		$sql = "SELECT ID FROM " . $tablename . " WHERE OrgID = '" . $orgid . "' AND VersionID = '$qid'";
//		$sql = "SELECT ID FROM tmp_ResponseFlat2 WHERE OrgID = '" . $orgid . "' AND VersionID = '$qid'";

		$result = $this->DB->query($sql);
		if ($line = $this->DB->fetch_array($result))
		{
				// survey was found
				$return_value = $line['ID'];
		}

		return $return_value;
	}

	// returns the User ID
	public function GetUserID()
	{
		return $this->UserID;
	}

	// returns the GUID
	public function GetGUID()
	{
		return $this->GUID;
	}

	// returns the name
	public function GetName()
	{
		return $this->Name;
	}

	// returns the password
	public function GetPassword()
	{
		return $this->Password;
	}

	// returns the Username
	public function GetUsername()
	{
		return $this->Username;
	}

	// returns the e-mail address
	public function GetEmail()
	{
		return $this->Email;
	}

	// returns the surveyor
	public function GetSurveyor()
	{
		return $this->Surveyor;
	}

	// returns the phone
	public function GetPhone()
	{
		return $this->Phone;
	}

	// returns the date created
	public function GetCreatedDate()
	{
		return $this->CreatedDate;
	}

	// returns the last login date
	public function GetLastLoginDate()
	{
		return $this->LastLoginDate;
	}

	// returns the survey completion date
	public function GetSurveyCompletionDate()
	{
		return $this->SurveyCompletionDate;
	}

	// returns whether or not the user has filled out their profile
	public function GetUserSaved()
	{
		if ($this->UserSaved == 'yes') return true;
		return false;
	}

	// returns the survey completion date
	public function GetPreviousLoginDate()
	{
		return $this->PreviousLoginDate;
	}

	// returns the survey completion date
	public function GetLoginCount()
	{
		return $this->LoginCount;
	}

	// returns areas of expertise
	public function GetEXPERTISE()
	{
		return $this->EXPERTISE;
	}

	// returns markets
	public function GetMARKETS()
	{
		return $this->MARKETS;
	}
}
?>