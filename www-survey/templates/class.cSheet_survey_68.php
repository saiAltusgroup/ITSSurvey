<?php
// DK - May 4, 2009

class cSheet_survey_68 extends cSheetITS
{
	private	$name = 'VO67';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What is the economic face rate and development yield for a Class "A" suburban building? <strong><em>(Development yield is defined as face rate total costs including hard and soft costs and land).</em></strong></p><p><strong><em>Please use decimals where possible</em></strong></p></div>');
		$this->SetFooter('<div class="descr"><p><em>Development Yield: As a percentage of total development cost</em></p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arr4, $this->market);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Suburban "A" Economic Face Rate &amp; Development Yield';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($this->name, $this->arr4, $this->market, $this->market_example_15);
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
