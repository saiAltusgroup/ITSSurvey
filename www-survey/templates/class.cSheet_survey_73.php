<?php
// DK - May 4, 2009

class cSheet_survey_73 extends cSheetITS
{
	private	$name = 'VR33';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Single Tenant Retail Building</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark single tenant retail building is defined as:</p><ul><li>Greater than 80,000 square feet</li><li>Of recent vintage (+/- 5 years old)</li><li>Suburban location, prominent arterial position</li><li>25% site coverage with good parking (minimum 5.0 per 1,000 s.f.)</li><li>15-year lease</li><li>Years one to five at market with market contractual rental growth in years six and eleven</li><li>Investment grade covenant strength (e.g. Home Depot, RONA and Canadian Tire)</li><li>For the purpose of this survey, assume the locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrRate_1, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Single-Tenant Retail Building';
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($this->name, $this->arrRate_1, $this->market, null);
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'D_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
