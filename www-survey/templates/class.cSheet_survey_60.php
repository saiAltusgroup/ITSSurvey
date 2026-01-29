<?php
// DK - May 4, 2009

class cSheet_survey_60 extends cSheetITS
{
	private	$name = 'MI50';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>If your answer to the previous question was <strong>Yes</strong>, what is the total sales cost that you utilize?</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arr8, $this->salescostaspercentsaleprice);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Total Sales At Reversion';
		$this->head_type='Property Sale Price';
		$this->head_example='';
		$this->DrawInputMatrix($this->name, $this->arr8, $this->salescostaspercentsaleprice, null);
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'A');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'E');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
