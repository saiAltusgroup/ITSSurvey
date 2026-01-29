<?php
// DK - Oct 22, 2008

class cSheet_survey_19 extends cSheetITS
{
	private	$name = 'VR60';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Vacancy &amp; Credit Allowance</u></strong><br />What is the vacancy and credit allowance utilized for the following retail property types?</p></div>');

//		$this->AddInput(new cInputTextBox('s19_60VPR', 'Retail Vacancy & Credit Allowance', ''));

		$this->AddInputMatrix_VacancyAndCredit($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 0, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE);
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'A_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'E_');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_VacancyAndCredit($this->name);
	}
}

?>
