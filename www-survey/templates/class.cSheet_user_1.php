<?php
// DK - Oct 22, 2008

class cSheet_user_1 extends cSheet
{
	public function /* void */ Create()
	{
		$this->AddInput(new cInputTextBox('CONTRIB', 'Client Contributor', ''));
		$this->FindInput('CONTRIB')->SetSize(60);
		$this->FindInput('CONTRIB')->SetMaxlength(100);

		$this->AddInput(new cInputTextBox('COMPANY', 'Client Company', ''));
		$this->FindInput('COMPANY')->SetSize(60);
		$this->FindInput('COMPANY')->SetMaxlength(100);

		$this->AddInput(new cInputTextBox('PHONE', 'Client Phone Number', ''));
		$this->FindInput('PHONE')->SetSize(60);
		$this->FindInput('PHONE')->SetMaxlength(30);

		$this->AddInput(new cInputTextBox('FAX', 'Client Fax Number', ''));
		$this->FindInput('FAX')->SetSize(60);
		$this->FindInput('FAX')->SetMaxlength(30);

		$this->AddInput(new cInputTextBox('SURVEYOR', 'Altus Group Surveyor', ''));
		$this->FindInput('SURVEYOR')->SetSize(60);
		$this->FindInput('SURVEYOR')->SetMaxlength(100);

		$this->AddInput(new cInputTextBox('Email', 'Email Address', ''));
		$this->FindInput('Email')->SetSize(60);
		$this->FindInput('Email')->SetMaxlength(100);

		$this->AddInput(new cInputChoiceSingleRadiolist('INDUSTRY', 'Industry', '', cAcquirer::AcquireIndustry()));
		$this->FindInput('INDUSTRY')->SetDescription('Select one of the following:');

		$this->AddInput(new cInputChoiceMultiChecklist('EXPERTISE', 'Areas of Expertise', '', cAcquirer::AcquireExpertise()));
		$this->FindInput('EXPERTISE')->SetDescription('Select all that apply:');

		$this->AddInput(new cInputChoiceMultiChecklist('MARKETS', 'Markets', '', cAcquirer::AcquireMarkets()));
		$this->FindInput('MARKETS')->SetDescription('Select all that apply:');

		$this->AddInput(new cInputLiteral('a', 'b', 'Please choose a user name and password. If you are currently a subscriber to the Investment Trends Survey online results, you may use the same user name and password.'));

		$this->AddInput(new cInputTextBox('Username', 'User Name', ''));
		$this->FindInput('Username')->SetSize(30);
		$this->FindInput('Username')->SetMaxlength(30);

		$this->AddInput(new cInputTextPassword('Password', 'Password', ''));
		$this->FindInput('Password')->SetSize(30);
		$this->FindInput('Password')->SetMaxlength(30);
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$validator = new cValidator(true);

		$ret &= $validator->ValidateStringLength($this->FindInput('CONTRIB'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);
		$ret &= $validator->ValidateStringLength($this->FindInput('COMPANY'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);
		$ret &= $validator->ValidateStringLength($this->FindInput('PHONE'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);
		$ret &= $validator->ValidateStringLength($this->FindInput('SURVEYOR'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);

		$ret &= $validator->ValidateStringLength($this->FindInput('INDUSTRY'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);
		$ret &= $validator->ValidateStringLength($this->FindInput('EXPERTISE'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);
		$ret &= $validator->ValidateStringLength($this->FindInput('MARKETS'), 'gt', 0, _ERROR_MUST_NOT_BE_EMPTY);

		$ret &= $validator->ValidateEmail($this->FindInput('Email'), _ERROR_MUST_BE_EMAIL);

		$ret &= $validator->ValidateStringLength($this->FindInput('Username'), 'gt', 2, _ERROR_MUST_BE_MORETHAN3CHARS);
		$ret &= $validator->ValidateStringLength($this->FindInput('Password'), 'gt', 2, _ERROR_MUST_BE_MORETHAN3CHARS);

		// checks for valid postal code
		//$ret &= $validator->ValidatePostalCode($this->FindInput('POSTAL'), _ERROR_INVALID_POSTAL);

		return $ret;
	}
}
?>