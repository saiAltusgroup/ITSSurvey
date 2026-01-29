<?php
// BD - Feb 23, 2009

class cSheet_survey_36 extends cSheetITS
{
	private	$name = 'VO123';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What percentage change (positive or negative) in market Face Rates do you expect for each of the following markets?</p></div>');
		$this->AddInputMatrix_TermMarket($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_Arr3_B($this->name, 'Suburban Class “B” Face Rate Change<br />% Change (Positive or Negative)');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			//if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 0, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE);

			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'A_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'B_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'C_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'D_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
