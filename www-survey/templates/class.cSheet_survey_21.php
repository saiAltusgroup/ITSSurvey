<?php
// DK - Oct 22, 2008

class cSheet_survey_21 extends cSheetITS
{
	private	$name = 'VI33';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Multi-Tenant Industrial Building</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark multi-tenant industrial building is defined as:</p><ul><li>60 - 100,000 s.f. with 15% to 25% office finish and minimum 18 ft. clear height (smaller in some markets)</li><li>Of recent vintage (+/- 10 years old)</li><li>Unit size range from 3,000 to 8,000 with 15 to 25 tenants</li><li>Suburban location</li><li>40% site coverage with good turning radius, one shipping door per demised unit and good parking</li><li>Year one contract at <strong>MARKET NET FACE RATES</strong> with annual market rental growth</li><li>Mostly local covenants but may have some recognizable nationals</li><li>Stabilized occupancy and no month-to-month tenancies</li><li>Lease rollover of approximately 20%-30% per annum for the next three years</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s20_33VPI_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI)', ''));
//		$this->AddInput(new cInputTextBox('s20_33VPI_B', 'b. $ Per s.f.', ''));
//		$this->AddInput(new cInputTextBox('s20_33VPI_C', 'c. Internal Rate of Return', ''));
//		$this->AddInput(new cInputTextBox('s20_33VPI_D', 'd. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s20_33VPI_E', 'e. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));

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
//		parent::ParentDisplay();

		$this->DrawInputMatrix_MultiIndustrial($this->name);
	}
}

?>
