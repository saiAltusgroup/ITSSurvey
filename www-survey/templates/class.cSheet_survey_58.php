<?php
// DK - May 4, 2009

class cSheet_survey_58 extends cSheetITS
{
	private	$name = 'VM40';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Development Yield &amp; Stabilized Cap Rate For New Construction</u></strong><br />What is the expected development yield required to stimulate new rental building construction? Development yield is defined as fully occupied Net Operating Income (NOI) ? total costs (i.e. land, hard and soft). Secondly, once the development is completed, what is the expected stabilized cap rate? (Central location with strong amenities and comparable to an attractive condominium position.)</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arr6, $this->market);

		$this->SetFooter('<div class="descr"><p><sup>1</sup> Stabilized Cap rate based on fully leased and developed net income stream for the new building to be sold to a third party.</p></div>');

	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Multi-Unit Residential Development Yield &amp; Cap Rate';
		$this->head_type='Type';
		$this->DrawInputMatrix($this->name, $this->arr6, $this->market, '');
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
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
