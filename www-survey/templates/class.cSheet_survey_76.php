<?php
// DK - May 4, 2009

class cSheet_survey_76 extends cSheetITS
{
	private	$name = 'CD82';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What three companies do you feel are the most active sources of first mortgage debt capital?</p></div>');

		$this->AddInput(new cInputChoiceSingleSelectlist($this->name . 'CD1', '1st Source: ', '', cAcquirer::AcquireMortgageCompanies()));
		$this->AddInput(new cInputTextBox($this->name . 'CD1_other', 'If other, please specify:', ''));
		$this->AddInput(new cInputChoiceSingleSelectlist($this->name . 'CD2', '2nd Source: ', '', cAcquirer::AcquireMortgageCompanies()));
		$this->AddInput(new cInputTextBox($this->name . 'CD2_other', 'If other, please specify:', ''));
		$this->AddInput(new cInputChoiceSingleSelectlist($this->name . 'CD3', '3rd Source: ', '', cAcquirer::AcquireMortgageCompanies()));
		$this->AddInput(new cInputTextBox($this->name . 'CD3_other', 'If other, please specify:', ''));
	}


	public function /* cBool */ Validate()
	{
//		return $this->Validate_Ignored(true);
		return true;
	}

	public function /* void */ CustomDisplay()
	{
		parent::ParentDisplay();
	}
}

?>
