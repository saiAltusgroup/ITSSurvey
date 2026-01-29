<?php
// DK - Oct 22, 2008

class cSheet_survey_25 extends cSheetITS
{
	private	$name = 'VM118';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What is the annual tenant turnover (in percentage terms) for the <strong>Benchmark Suburban Multi-Unit Residential</strong> rental building?</p></div>');

//		$this->AddInput(new cInputTextBox('s25_118VPMR', '% Turnover', ''));

		$this->AddInputMatrix_SuburbanMultiAnnual($this->name, 'Value');

	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 1, 100, _ERROR_MUST_BE_RATE);
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_SuburbanMultiAnnual($this->name);
	}
}

?>
