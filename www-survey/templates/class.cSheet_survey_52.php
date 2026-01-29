<?php
// DK - May 4, 2009

class cSheet_survey_52 extends cSheetITS
{
	private	$name = 'VO136';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What year one yield (overall capitalization rate) would you consider appropriate for a ground lease for the leasehold <strong>Benchmark Downtown Class “AA” office building</strong>, defined below:</p><ul><li>Un-expired ground lease term of 80 years</li><li>Current rent is fixed for the next five years at market</li><li>No participation ground rent</li><li>No option to acquire the leased fee interest</li><li>The subsequent renewals are for ten-year terms</li><li>The renewal rent is based on the market value of the land unimproved, multiplied by the market rate of return</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have these leasing characteristics</li></ul></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrRate_3, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title= 'Ground Lease Downtown Class "AA" Office Building';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arrRate_3, $this->market, $this->market_example_13);
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
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'B_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
