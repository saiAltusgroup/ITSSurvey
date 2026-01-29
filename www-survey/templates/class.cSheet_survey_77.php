<?php
// DK - May 4, 2009

class cSheet_survey_77 extends cSheetITS
{
	private	$name = 'MI96';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What is the marketing time for the following property types? Marketing time is defined as the time expired between the listing date and the closing date expressed in number of months.</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrClasses_1, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Marketing Time<br />No. of months';
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($this->name, $this->arrClasses_1, $this->market, null);
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
