<?php
// DK - Oct 22, 2008

class cSheet_survey_16 extends cSheetITS
{
	private	$name = 'VR28';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Community Mall</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark Community Mall is defined as:</p><ul><li>150,000 to 250,000 s.f. enclosed (maybe smaller in some markets)</li><li>Intersection of two major arterials</li><li>Mix of national chains and local independents</li><li>Anchored by one major chain supermarket and at least one department / discount store</li><li>CRU leased at <strong>MARKET RATES</strong></li><li>CRU sales currently average less than $300 per s.f.</li><li>Typically of older construction, approximately 20 years old, lying in mature area</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul></div>');

//		$this->AddInput(new cInputTextBox('s16_28VPR_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI)', ''));
//		$this->AddInput(new cInputTextBox('s16_28VPR_B', 'b. Internal Rate of Return', ''));
//		$this->AddInput(new cInputTextBox('s16_28VPR_C', 'c. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s16_28VPR_D', 'd. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));

		$this->AddInputMatrix_arrRatemarket($this->name, 'Value');
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

		$this->DrawInputMatrix_CommunityMall($this->name);
	}
}

?>
