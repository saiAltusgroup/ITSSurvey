<?php
// BD - Feb 23, 2009

class cSheet_survey_32 extends cSheetITS
{
	private	$name = 'VO55';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>LEASEHOLD Benchmark Downtown Class "AA" </strong> office property, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The ground lease for the leasehold Downtown Class "AA" building is defined as:</p><ul><li>Un-expired ground lease term of 80 years.</li><li>No participation ground rent</li><li>No option to acquire the leased fee interest</li><li>Current rent is fixed for the next five years</li><li>The subsequent renewals are for ten-year terms</li><li>The renewal rent is based on the market value of the land unimproved, multiplied by the market rate of return</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul></div>');
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
		$this->DrawInputMatrix_Arr1_B($this->name, 'Leasehold Downtown Class "AA" Office Building', 1);
	}
}

?>
