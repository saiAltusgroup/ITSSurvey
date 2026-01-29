<?php
// DK - May 4, 2009

class cSheet_survey_54 extends cSheetITS
{
	private	$name = 'VO134';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Net Effective / Face Rates</u></strong><br /><strong><em>For New Tenants leasing space for a five year term</em></strong>, what is your opinion of the current <strong><em>market</em></strong> Net Effective Rate (base building condition <sup>(2)</sup>) and Face Rate (built-out from base building condition) in the each of the following markets?</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arr2, $this->market);

		$this->SetFooter('<div class="descr"><p><em>Net Effective Rate - After Realty Commissions and Discounted at 10% monthly</em><br/><em>Face Rate - Sum of Net Effective Rate and Amortized Inducements and Commissions</em><br/><em><sup>(1)</sup> Mid-rise and City view</em><br/><em><sup>(2)</sup> Drywall perimeter walls, finished ceiling system (i.e. acoustic tile, HVAC troughers and fluorescent lighting fixtures, base concrete floor. No interior demising walls.</em></p></div>');
	
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Downtown Class “B” Net Effective / Face Rates<br />New Tenants - 5 Year Terms';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arr2, $this->market, $this->market_example_14);
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
