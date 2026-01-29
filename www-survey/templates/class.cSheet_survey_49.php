<?php
// BD - Feb 24, 2009

class cSheet_survey_49 extends cSheetITS
{
	private	$name = 'CD83';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What is the best available combination of lending parameters for conventional first mortgage debt for the indicated asset classes?</p><p><strong><em>Only indicate answers to the asset classes in which you, as a lender, are interested in competing for financing or, as a borrower, were recently involved in financing your realty asset(s).</em></strong></p><p><strong><em>Please use decimals where possible</em></strong></p></div>');

		$this->AddInputMatrix_CapitalBenchmark($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_CapitalBenchmark($this->name, 'Conventional Debt Market Parameters');
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
			//$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B_');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'G_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
