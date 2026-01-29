<?php
// BD - Feb 23, 2009

class cSheet_survey_41 extends cSheetITS
{
	private	$name = 'VR70';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Food Anchored Retail Strip</strong> centres, please provide your opinion of current <strong>"MARKET"</strong> valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark food anchored retail strip is defined as:</p><ul><li>10 to 15 years old</li><li>75,000 to 125,000 s.f.</li><li>Intersection of two arterials</li><li>Secure local position</li><li>Mature residential district</li><li>Major chain supermarket trading at &ge;$450 per s.f. with &ge;10 year lease</li><li>CRU mix contains blend of national / regional chains and local / independents</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><strong><em>Please use decimals where possible</em></strong></div>');

		$this->AddInputMatrix_arrRatemarket_1($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_arrRatemarket_12($this->name, 'Food Anchored Retail Strip');
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
}

?>
