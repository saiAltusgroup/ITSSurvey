<?php
// DK - May 4, 2009

class cSheet_survey_65 extends cSheetITS
{
	private	$name = 'VO64';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Suburban Class "A" Office Building</strong>, please provide your opinion of <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The suburban Class "A" building is defined as:</p><ul><li>100,000 - 170,000 s.f.</li><li>Class "A" physical features - 5 to 10 years old</li><li>3.0 stalls per 1,000 s.f. surface parking</li><li>Prominent position within established and well recognized suburban business park</li><li>Multi tenant</li><li>5% vacancy</li><li>10% roll-over per year</li><li>Leased at market rates to good covenant tenants</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrRateSF_5, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Suburban Class "A" Office Building';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arrRateSF_5, $this->market, $this->market_example_15);
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
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
