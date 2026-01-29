<?php
// DK - Oct 22, 2008

class cSheet_survey_2 extends cSheetITS
{
	private	$name = 'VO99';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Downtown Class "AA" Office Building</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The Downtown Class "AA" office building is defined as:</p><ul><li>400,000 to 900,000 s.f. (may be smaller in some markets).</li><li>Class "AA" physical features</li><li>Strong location in local market</li><li>5% existing vacancy</li><li>10% roll-over per year</li><li>Leased at MARKET RATES to good covenant tenants</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		//$this->AddInput(new cInputTextBox('s2_99VPO_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI before leasing costs)', ''));
		//$this->AddInput(new cInputTextBox('s2_99VPO_B', 'b. $ Per s.f.', ''));
		//$this->AddInput(new cInputTextBox('s2_99VPO_C', 'c. Internal Rate of Return', ''));
		//$this->AddInput(new cInputTextBox('s2_99VPO_D', 'd. Terminal Cap Rate', ''));
		//$this->AddInput(new cInputTextBox('s2_99VPO_E', 'e. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));

		$this->AddInputMatrix_arrRateSF($this->name, 'Value');
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

	public function /* void */ CustomDisplay()
	{
		//parent::ParentDisplay();

		$this->DrawInputMatrix_Arr1($this->name);
	}
}

?>
