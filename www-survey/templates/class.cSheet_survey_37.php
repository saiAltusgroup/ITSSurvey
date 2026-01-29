<?php
// BD - Feb 23, 2009

class cSheet_survey_37 extends cSheetITS
{
	private	$name = 'VR61';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What Annual percent increase (or decrease) do you currently apply to Tier I Regional Mall market rents, expenses and sales?</p></div>');
		$this->AddInputMatrix_arrRateSF_2($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_Arr6_2($this->name, 'Tier I Regional Mall Rent, Expense and Sales Annual Percent Change');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			//if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 0, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE);
			
			
			
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'A_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'B_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'C_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'D_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'E_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
