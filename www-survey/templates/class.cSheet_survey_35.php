<?php
// BD - Feb 23, 2009

class cSheet_survey_35 extends cSheetITS
{
	private	$name = 'VO122';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For <strong><em>New Tenants leasing space for five year terms</em></strong>, what is the current <strong><em>market</em></strong> Net Effective Rate (base building condition <sup>(1)</sup>) and Face Rate (built-out from base building condition) in the each of the following markets?</p></div>');
		$this->AddInputMatrix_Arr2($this->name, 'Value');
		$this->SetFooter('<div class="descr"><p><em>Net Effective Rate - After Realty Commissions and Discounted at 10% monthly</em><br/><em>Face Rate - Addition of Net Effective Rate and Amortized Inducements and Commissions</em><br/><em><sup>(1)</sup> Drywall perimeter walls, finished ceiling system (i.e. acoustic tile, HVAC troughers and fluorescent lighting fixtures, base concrete floor. No interior demising walls.</em></p></div>');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_Arr2_B($this->name, 'Suburban Class “B” Net Effective/Face Rates<br />New Tenants - 5 Year Terms');
//		$inputLit = $this->FindInput('description');
//		$this->DisplayInput($inputLit, false);
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
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
