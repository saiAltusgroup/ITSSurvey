<?php
// DK - May 4, 2009

class cSheet_survey_74 extends cSheetITS
{
	private	$name = 'VM81';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark <u>Downtown</u> Multi-Unit Residential Building</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months <!--<strong>with special reference to the current direction of debt costs</strong>-->.<br />The benchmark downtown multi-unit residential building is defined as:</p><ul><li>150-300 Suites - high-rise</li><li>1960s construction</li><li>Well maintained with no deferred maintenance</li><li>Good downtown location</li><li>Good proximity to retail and transportation facilities</li><li>Market rents with stabilized operating expenses</li><li>Average of 30%-35% annual tenant turnover</li><li>Market level financing and loan to value ratio</li><li>For the purpose of this survey, assume the buildings/locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><p><strong><em>Please use decimals where possible. <u>When entering $ Price Per Suite, please provide the full dollar value (i.e. include all the zeroes), not a value in thousands of dollars.</u></em></strong></p></div>');
		$this->SetFooter('<div class="descr"><p><em>Based on typical LTV of 75.00% of value at market rates. No second mortgage.</em></p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrMultiUnitResidentialRate, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Downtown Multi-Unit Residential';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arrMultiUnitResidentialRate, $this->market, $this->market_example_16);
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D_');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'E_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
