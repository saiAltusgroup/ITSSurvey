<?php
// BD - Feb 24, 2009

class cSheet_survey_44 extends cSheetITS
{
	private	$name = 'VI128';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For New Tenants Leasing Space in a Benchmark Single Tenant Industrial building for a ten-year term, what is the current <strong><em>market net (face) rental rate</em></strong> in the each of the following markets?<br /><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix_arr_rentalrate($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
	$this->DrawInputMatrix_arr_rentalrate($this->name, 'Net (Face) Rental Rate - Single Tenant Industrial<br />New Tenants - 10 Year Terms');
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
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
