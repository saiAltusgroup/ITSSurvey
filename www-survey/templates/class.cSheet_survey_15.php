<?php
// DK - Oct 22, 2008

class cSheet_survey_15 extends cSheetITS
{
	private	$name = 'VR102';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Tier II Regional Mall Rent &amp; Expense Annual Percent Change</u></strong><br />What Annual percent increase (or decrease) do you currently apply to <strong>Tier II Regional Mall</strong> market rents and expenses?</p></div>');

//		$this->AddInput(new cInputTextBox('s15_102VPR_A', 'Market Rent Inflation', ''));
//		$this->AddInput(new cInputTextBox('s15_102VPR_B', 'Expenses Inflation', ''));

		$this->AddInputMatrix_Tier2RegionalMallInflation($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'A_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'B_');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_Tier2RegionalMallInflation($this->name);
	}
}

?>
