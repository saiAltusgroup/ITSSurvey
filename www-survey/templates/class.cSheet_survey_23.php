<?php
// DK - Oct 22, 2008

class cSheet_survey_23 extends cSheetITS
{
	private	$name = 'VI134';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Net Face Rental Rate Change</u></strong><br />What percentage change (positive or negative) in market net rental rates do you expect for each of the following markets?</p></div>');

//		$this->AddInput(new cInputTextBox('s23_134VPI_A', 'a. Next 12 Months', ''));
//		$this->AddInput(new cInputTextBox('s23_134VPI_B', 'b. 13 - 24 Months', ''));
//		$this->AddInput(new cInputTextBox('s23_134VPI_C', 'c. 25 - 36 Months', ''));
//		$this->AddInput(new cInputTextBox('s23_134VPI_D', 'd. Stabilized Long Term', ''));

		$this->AddInputMatrix_TermMarket($this->name, 'Value');
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

		$this->DrawInputMatrix_NetFaceGrowth($this->name);
	}
}

?>
