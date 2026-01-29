<?php
// DK - Oct 22, 2008

class cSheet_survey_29 extends cSheetITS
{
	private	$name = 'CC48';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Real Estate Equity Price Performance</u></strong><br />Excluding the potential for take-overs, how will the price of Real Estate Equities perform over the next 12 months? <strong>Indicate your response by inserting a percentage increase or decrease in the appropriate box below.</strong></p></div>');

//		$this->AddInput(new cInputTextBox('s29_48CSCM_A', 'Large Cap REITs (% increase or decrease)', ''));
//		$this->AddInput(new cInputTextBox('s29_48CSCM_B', 'Small Cap REITs (% increase or decrease)', ''));
//		$this->AddInput(new cInputTextBox('s29_48CSCM_C', 'Small Cap Corps (% increase or decrease)', ''));

		$this->AddInputMatrix_PricePerformance($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, -100, 100, _ERROR_MUST_BE_IN_NUMERIC_RANGE4);
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_PricePerformance($this->name);
	}
}

?>
