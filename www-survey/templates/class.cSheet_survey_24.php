<?php
// DK - Oct 22, 2008

class cSheet_survey_24 extends cSheetITS
{
	private	$name = 'VM36';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark <u>Suburban</u> Multi-Unit Residential Building</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark suburban multi-unit residential building is defined as:</p><ul><li>150 - 300 Suites – high-rise (may be smaller in smaller markets)</li><li>Early to mid 1970\'s construction</li><li>Well maintained with no deferred maintenance</li><li>Good suburban location</li><li>Good proximity to retail and transportation facilities</li><li>Market rents with stabilized operating expenses</li><li>Average of 30%-35% annual tenant turnover</li><li>25% cash to first mortgage at current rates</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><p><strong><em>Please use decimals where possible. <u>When entering $ Price Per Suite, please provide the full dollar value (i.e. include all the zeroes), not a value in thousands of dollars.</u></em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s24_36VPMR_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI)', ''));
//		$this->AddInput(new cInputTextBox('s24_36VPMR_B', 'b. $ Price Per Suite', ''));
//		$this->AddInput(new cInputTextBox('s24_36VPMR_C', 'c. Internal Rate of Return (IRR)', ''));
//		$this->AddInput(new cInputTextBox('s24_36VPMR_D', 'd. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s24_36VPMR_E', 'e. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));

		$this->AddInputMatrix_SuburbanMulti($this->name, 'Value');

		$this->SetFooter('<div class="descr"><p><em>* Based on typical LTV of 75.00% of value at market rates. No second mortgage.</em></p><p><em>** Property fully meeting our benchmark guidelines does not currently exist within this market. In order to gauge market opinion please assume property has parameters listed above.</em></p></div>');
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

			// minimum 5 digits as requested by Kevin
			$ret &= $this->Validate_Large($input, _ERROR_MUST_BE_LARGE, 'B_');

			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'A_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D_');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'E_');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_SuburbanMulti($this->name);
	}
}
?>
