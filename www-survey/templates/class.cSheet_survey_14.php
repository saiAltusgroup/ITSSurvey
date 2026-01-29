<?php
// DK - Oct 22, 2008

class cSheet_survey_14 extends cSheetITS
{
	private	$name = 'VR101';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Tier II Regional Mall</strong> in a secondary market, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark Tier II regional mall is defined as:</p><ul><li>400,000 s.f. - 600,000 s.f. (maybe smaller in some markets)</li><li>Dominant only in a local sense</li><li>Renovated in the past 10 years</li><li>2 anchor tenants occupy at least 40% of rentable area</li><li>CRU leased at <strong>MARKET RATES</strong></li><li>CRU sales are relatively flat and are &lt; $400 per s.f.</li><li>Expansion potential not to be considered</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul></div>');

//		$this->AddInput(new cInputTextBox('s14_101VPR_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI)', ''));
//		$this->AddInput(new cInputTextBox('s14_101VPR_B', 'b. Internal Rate of Return', ''));
//		$this->AddInput(new cInputTextBox('s14_101VPR_C', 'c. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s14_101VPR_D', 'd. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));

		$this->AddInputMatrix_Tier2RegionalMall($this->name, 'Value');
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

		$this->DrawInputMatrix_Tier2RegionalMall($this->name);
	}
}

?>
