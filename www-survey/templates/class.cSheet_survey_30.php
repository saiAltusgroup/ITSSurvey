<?php
// DK - Oct 22, 2008

class cSheet_survey_30 extends cSheetITS
{
	private	$name = 'CC43';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What is your Total Required Rate of Return, dividend plus appreciation, for Real Estate Equities over the next 12 months? <strong>Indicate your response by inserting the percentage return in the appropriate box below.</strong></p></div>');

//		$this->AddInput(new cInputTextBox('s30_43CSCM_A', 'REITs - Return (%)', ''));
//		$this->AddInput(new cInputTextBox('s30_43CSCM_B', 'Publicly Traded Corporations - Return (%)', ''));

		$this->AddInputMatrix_RequiredReturn($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			//lacfre 2022-06
			//if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 1, 100, _ERROR_MUST_BE_RATE);
			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, -100, 100, _ERROR_MUST_BE_IN_NUMERIC_RANGE4); 
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_RequiredReturn($this->name);
	}
}

?>
