<?php
// DK - Oct 22, 2008

class cSheet_survey_7 extends cSheetITS
{
	private	$name = 'VO119';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark <u>Midtown</u> Class "A" Office Building in <u>Montreal</u></strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark Midtown Class "A" Office Building is defined as:</p><ul><li>50,000 to 100,000 square feet</li><li>5 to 15 years old</li><li>Class "A" physical features</li><li>Combination of indoor/outdoor parking; 1 to 2 stalls per 1,000 square feet</li><li>Good location</li><li>Montreal Urban Community location, outside downtown but within 10 km from Place Ville-Marie</li><li>Multi-tenant</li><li>Vacancy at market average</li><li>10% annual rollover</li><li>Leased at <strong>MARKET RATES</strong> to good covenant tenants</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s7_119VPO_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI before leasing costs)', ''));
//		$this->AddInput(new cInputTextBox('s7_119VPO_B', 'b. $ Per s.f.', ''));
//		$this->AddInput(new cInputTextBox('s7_119VPO_C', 'c. Internal Rate of Return', ''));
//		$this->AddInput(new cInputTextBox('s7_119VPO_D', 'd. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s7_119VPO_E', 'e. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3', ''));

		$this->AddInputMatrix_Arr6($this->name, 'Value');
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

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_Arr6($this->name);
	}
}

?>
