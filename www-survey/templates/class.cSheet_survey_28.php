<?php
// DK - Oct 22, 2008

class cSheet_survey_28 extends cSheetITS
{
	private	$name = 'CC47';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Real Estate Equity Availability</u></strong><br />Describe the availability of New Real Estate Equity for large and small cap real estate companies.</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'A_R', 'Large Cap REITs', '', cAcquirer::AcquireRating()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B_R', 'Small Cap REITs', '', cAcquirer::AcquireRating()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B_C', 'Small Cap Corps', '', cAcquirer::AcquireRating()));
	}

	public function /* cBool */ Validate()
	{
		return $this->Validate_Ignored(true);
	}

	public function /* void */ CustomDisplay()
	{
		parent::ParentDisplay();
	}
}

?>
