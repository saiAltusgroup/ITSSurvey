<?php
// BD - March 03, 2009, based on SDC password retrieval

class cSheet_password_1 extends cSheet
{
	public function /* void */ Create()
	{
		$this->AddInput(new cInputLiteral('login_instructions', '', _ENTER_EMAIL));

		$this->AddInput(new cInputTextBox('email', _EMAIL, ''));
		$this->FindInput('email')->SetSize(30);
		$this->FindInput('email')->SetMaxlength(255);
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$validator = new cValidator(true);
		// checks to make sure q1 is not empty
		$ret &= $validator->ValidateStringLength($this->FindInput('email'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);
		return $ret;
	}
}
?>