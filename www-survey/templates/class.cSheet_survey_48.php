<?php
// BD - Feb 24, 2009

class cSheet_survey_48 extends cSheetITS
{
	private	$name = 'VM127';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark <u>Midtown</u> Multi-Unit Residential Building in <u>Montreal</u></strong>, please provide your opinion of current "<strong>MARKET</strong>" valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The Benchmark Multi-Unit Residential Building is defined as:</p><ul><li>200 - 300 suites, high-rise</li><li>25 years old</li><li>Well maintained with no deferred maintenance</li><li>Good midtown location</li><li>Montreal Urban Community location, outside downtown but within 10 km from Place Ville-Marie</li><li>Good proximity to retail and transportation</li><li>Stabilized rents and stabilized operating expenses</li><li>Few, if any rental board cases</li><li>Average 15% to 25% annual rollover</li><li>25% cash to first mortgage at current rates</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix_Arr6_3($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_Arr6_3($this->name, 'Montreal Midtown Multi-Unit Residential Building');
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'E');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
