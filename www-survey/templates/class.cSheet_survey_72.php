<?php
// DK - May 4, 2009

class cSheet_survey_72 extends cSheetITS
{
	private	$name = 'VR139';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Power Centre - Vacant Land Value</u></strong><br />What is the $ / per acre value of vacant land (serviced to the lot line) of a site suitable for development of a benchmark Power Centre?</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->peracre, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Power Centre - Vacant Land Value';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->peracre, $this->market, $this->market_example_11);
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
