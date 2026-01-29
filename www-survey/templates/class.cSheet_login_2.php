<?php
// DK - Oct 22, 2008

class cSheet_login_2 extends cSheet
{
	public function /* void */ Create()
	{
		$this->AddInput(new cInputLiteral('login_instructions', '', 'Please enter your admin username and password. If you have lost or misplaced this username and/or password, please contact BOLD.'));

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

		// checks to make sure username is correct
		$ret &= $validator->ValidateStringEquals($this->FindInput('username'), 'itsadmin', false, '');

		// checks to make sure password is correct
		//$ret &= $validator->ValidateStringEquals($this->FindInput('password'), '1element!Q#4', true, '');
		$ret &= $validator->ValidateStringEquals($this->FindInput('password'), 'XNb4R#3!Je34', true, '');


		return $ret;
	}
}
?>