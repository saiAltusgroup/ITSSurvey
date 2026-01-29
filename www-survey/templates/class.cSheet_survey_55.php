<?php
// DK - May 4, 2009

class cSheet_survey_55 extends cSheetITS
{
	private	$name = 'VO135';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Face Rate Change</u></strong><br />What percentage change (positive or negative) in market Face Rates do you expect for each of the following markets?</p></div>');

		$this->AddInputMatrix_TermMarket($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_Arr3_D($this->name, 'Downtown Class "B" Face Rate Change<br />% Change (positive or negative)');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, -1000000, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE3);
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'A');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'B');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'C');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'D');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
