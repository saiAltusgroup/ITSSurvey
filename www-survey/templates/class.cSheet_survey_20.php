<?php
// DK - Oct 22, 2008

class cSheet_survey_20 extends cSheetITS
{
	private	$name = 'VI75';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Single Tenant Industrial Building</strong>, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark single-tenant industrial building is defined as:</p><ul><li>100,000 s.f. single tenant warehouse property (maybe slightly smaller in some markets)</li><li>Pre-cast construction with 10% office finish and 28 ft. clear height</li><li>Of recent vintage (+/- 5 years old)</li><li>Suburban location, with good highway access</li><li>45% site coverage with good turning radius and good parking</li><li>10 to 15 shipping doors</li><li>10 year lease</li><li>Years one to five leased at <strong>MARKET RATE</strong> with market contractual rental growth in year six</li><li>Investment grade covenant strength (i.e. Standard &amp; Poor BBB or better)</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristic</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s20_75VPI_A', 'a. Overall Cap Rate (Yr 1 Stabilized NOI)', ''));
//		$this->AddInput(new cInputTextBox('s20_75VPI_B', 'b. $ Per s.f.', ''));
//		$this->AddInput(new cInputTextBox('s20_75VPI_C', 'c. Internal Rate of Return', ''));
//		$this->AddInput(new cInputTextBox('s20_75VPI_D', 'd. Terminal Cap Rate', ''));
//		$this->AddInput(new cInputTextBox('s20_75VPI_E', 'e. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No Change = 3)', ''));


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

		$this->DrawInputMatrix_SingleIndustrial($this->name);
	}
}

?>
