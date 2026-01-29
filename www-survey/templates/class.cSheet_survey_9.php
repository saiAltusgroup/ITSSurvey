<?php
// DK - Oct 22, 2008

class cSheet_survey_9 extends cSheetITS
{
	private	$name = 'VO121';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What percentage change (positive or negative) in market <strong>Face Rates</strong> do you expect for the <strong>Benchmark <u>Midtown</u> Class "A" Office Building in <u>Montreal</u></strong> over the following time periods?</p><p><strong><em>Please use decimals where possible</em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s9_121VPO_A', 'a. Next 12 Months', ''));
//		$this->AddInput(new cInputTextBox('s9_121VPO_B', 'b. 13 - 24 Months', ''));
//		$this->AddInput(new cInputTextBox('s9_121VPO_C', 'c. 25 - 36 Months', ''));
//		$this->AddInput(new cInputTextBox('s9_121VPO_D', 'd. Stabilized Long Term', ''));

		$this->AddInputMatrix_Arr8($this->name, 'Value');
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

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_Arr8($this->name);
	}
}

?>
