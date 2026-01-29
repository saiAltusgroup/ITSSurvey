<?php
// DK - Oct 22, 2008
// BD - Nov 05, 2008

class cSheet_survey_13 extends cSheetITS
{
	private	$name = 'VR100';

	public function /* void */ Create()
	{

		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Tier I Regional Mall</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark tier I regional mall is defined as:</p><ul><li>Greater than 700,000 sf (maybe smaller in some markets)</li><li>Dominant market presence</li><li>Renovated in the past 5 years (maybe older in certain markets)</li><li>2/3 anchor tenants occupy up to 50% of rentable area</li><li>CRU leased at <strong>MARKET RATES</strong></li><li>CRU sales have been growing and are in excess of &gt;$500 per s.f. sourced from 85% national or provincial tenants</li><li>Expansion potential not to be considered</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul></div>');
//		$this->AddInput(new cInputTextBox('s13_100VPR_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI)', ''));
//		$this->AddInput(new cInputTextBox('s13_100VPR_B', 'b. Internal Rate of Return', ''));
//		$this->AddInput(new cInputTextBox('s13_100VPR_C', 'c. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s13_100VPR_D', 'd. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));

		$this->head_title='Tier I Regional Mall';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->AddInputMatrix($this->name, 'Value', $this->arrRate, $this->market);

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

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

//		$this->DrawInputMatrix_Tier1RegionalMall('s13_100VPR');
		$this->DrawInputMatrix($this->name, $this->arrRate, $this->market, $this->market_example_3);
	}
}

?>
