<?php
// BD - Feb 23, 2009

class cSheet_survey_38 extends cSheetITS
{
	private	$name = 'VR69';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Power Centre</strong>, please provide your opinion of current "<strong>MARKET</strong>" valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The benchmark power centre is defined as:</p><ul><li>Dominant for the Region</li><li>&gt; 500,000 s.f.</li><li>5 to 10 years old</li><li>Major arterial position</li><li>Proximity to controlled highway access</li><li>Strong covenant, national tenants</li><li>10-15 year leases</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned leasing characteristics</li></ul><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix_arrRatemarket_1($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_arrRatemarket_1($this->name, 'Power Centre');
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
