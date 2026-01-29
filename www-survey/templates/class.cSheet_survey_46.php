<?php
// BD - Feb 24, 2009

class cSheet_survey_46 extends cSheetITS
{
	private	$name = 'VM38';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the benchmark suburban multi-unit residential rental building, what is the typical <strong>average budget</strong> for the following expenses?</p></div>');

		$this->AddInputMatrix_arrBudget($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_arrBudget($this->name, 'Suburban Multi-Unit Residential Rental Building Expenses');
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
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
