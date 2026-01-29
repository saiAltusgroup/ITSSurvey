<?php
// DK - Oct 22, 2008

class cSheet_survey_8 extends cSheetITS
{
	private	$name = 'VO120';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><em>For New Tenants leasing space for five year terms</em></strong>, what would the current <strong><em>market</em></strong> Net Effective Rate (base building condition <sup>(1)</sup>) and Face Rate (built-out from base building condition) be for the <strong>Benchmark <u>Midtown</u> Class "A" Office Building in <u>Montreal</u></strong>?</p><p><strong><em>Please use decimals where possible</em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s8_120VPO_A', 'a. Net Effective Rate (Base Building)', ''));
//		$this->AddInput(new cInputTextBox('s8_120VPO_B', 'b. Face Rate (Built-out from Base Building)', ''));

		$this->AddInputMatrix_Arr7($this->name, 'Value');

		$this->SetFooter('<div class="descr"><p><em>Net Effective Rate - After Realty Commissions and Discounted at 10% monthly</em><br/><em>Face Rate - Addition of Net Effective Rate and Amortized Inducements and Commissions</em><br/><em><sup>(1)</sup> Drywall perimeter walls, finished ceiling system (i.e. acoustic tile, HVAC troughers and fluorescent lighting fixtures, base concrete floor. No interior demising walls.</em></p></div>');
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'A');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_Arr7($this->name);
	}
}

?>
