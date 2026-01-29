<?php
// DK - May 4, 2009

class cSheet_survey_67 extends cSheetITS
{
	private	$name = 'VO66';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Face Rate Inflation</u></strong><br />What percentage change in market Face Rates do you expect for each of the following markets?</p><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrTerm, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Suburban Class "A" Face Rate Inflation<br />% Change';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arrTerm, $this->market, $this->market_example_15);
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, -1000000, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE);
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'A_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'B_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'C_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'D_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
