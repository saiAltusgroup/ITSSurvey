<?php
// BD - Feb 24, 2009

class cSheet_survey_47 extends cSheetITS
{
	private	$name = 'VM11';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark <u>Suburban</u> Multi-Unit Residential Building</strong> what is the <strong>leasehold Overall Capitalization Rate</strong>?</p></div>');

		$this->AddInputMatrix_arrRatemarket_2($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_arrRatemarket_2($this->name, 'Suburban Multi- Unit Residential Leasehold Parameters');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0)
			$ret &= $validator->ValidateNumericRange($input, 0, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE);
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, '');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
