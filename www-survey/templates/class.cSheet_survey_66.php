<?php
// DK - May 4, 2009

class cSheet_survey_66 extends cSheetITS
{
	private	$name = 'VO65';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Net Effective / Face Rates</u></strong><br /><strong><em>For New Tenants leasing space for five year terms,</em></strong> what is your opinion of the current <strong><em>market</em></strong> Net Effective Rate (base building condition<sup>(1)</sup>) and Face Rate (built-out from base building condition) in the each of the following markets?</p><p><strong><em>Please use decimals where possible</em></strong></p></div>');
		$this->SetFooter('<div class="descr"><p><em>Net Effective Rate - After Realty Commissions and Discounted at 10% monthly<br />Face Rate - Addition of Net Effective Rate and Amortized Inducements and Commissions<br />Vancouver: Mid-rise and City view<br /><sup>(1)</sup> Drywall perimeter walls, finished ceiling system (i.e. acoustic tile, HVAC troughers and flourescent lighting fixtures, base concrete floor. No interior demising walls.</em></p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arr2, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Suburban "A" Net Effective / Face Rates';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arr2, $this->market, $this->market_example_15);
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
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
