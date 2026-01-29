<?php
// DK - May 4, 2009

class cSheet_survey_56 extends cSheetITS
{
	private	$name = 'VR126';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Enclosed Community Mall Rent &amp; Expense Annual Percent Change</u></strong><br />What Annual percent increase (or decrease) do you currently apply to Enclosed Community Mall market rents and expenses?</p></div>');

		$this->AddInputPowerCentreRentExpense($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputPowerCentreRentExpense($this->name, 'Enclosed Community Mall Rent &amp; Expense Annual Percent Change');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'A_');
			$ret &= $this->Validate_Change($input, _ERROR_MUST_BE_CHANGE, 'B_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
