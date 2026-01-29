<?php
// DK - May 4, 2009

class cSheet_survey_53 extends cSheetITS
{
	private	$name = 'VO23';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Downtown Class "B" Office Building</strong>, please provide your opinion of current "<strong>MARKET</strong>" valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark Downtown Class "B" building is defined as:</p><ul><li>100,000 to 200,000 s.f. (maybe smaller in some markets)</li><li>Class "B" physical features - good condition</li><li>Limited on-site parking</li><li>Good location</li><li>Multi-tenant</li><li>5% vacant</li><li>10% roll-over per year</li><li>Leased at <strong>MARKET RATES</strong>. No triple "A" tenants</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrRateSF_4, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Downtown - Class "B" Office Building';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arrRateSF_4, $this->market, $this->market_example_14);
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
