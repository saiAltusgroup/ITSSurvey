<?php
// DK - Oct 22, 2008

class cSheet_login_1 extends cSheet
{
	public function /* void */ Create()
	{
		$this->AddInput(new cInputLiteral('login_instructions', '', 'Please enter the username and password you selected when setting up your Contributor Profile. If you have lost or misplaced your username and/or password, please follow the Support link on the main menu at the top of this page.'));

		$this->AddInput(new cInputTextBox('username', _USERNAME, ''));
		$this->FindInput('username')->SetSize(30);
		$this->FindInput('username')->SetMaxlength(255);
		$this->AddInput(new cInputTextPassword('password', _PASSWORD, ''));
		$this->FindInput('password')->SetSize(30);
		$this->FindInput('password')->SetMaxlength(255);
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$validator = new cValidator(true);

		// checks to make sure username is not empty
		$ret &= $validator->ValidateStringLength($this->FindInput('username'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);

		// checks to make sure password is not empty
		$ret &= $validator->ValidateStringLength($this->FindInput('password'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);

		return $ret;
	}
}
?>