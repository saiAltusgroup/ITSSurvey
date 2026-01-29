<?php
// DK - May 4, 2009

class cSheet_survey_70 extends cSheetITS
{
	private	$name = 'VO4';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For tenants that are expected to vacate on lease expiry, what is the office lag vacancy, (i.e. the number of months the space is forecast to remain vacant prior to releasing, inclusive of fixturing period), expressed in number of months?</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrClasses, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Office Lag Vacancy<br />No. of Months';
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($this->name, $this->arrClasses, $this->market, null);
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
