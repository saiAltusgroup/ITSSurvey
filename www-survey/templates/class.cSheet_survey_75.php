<?php
// DK - May 4, 2009

class cSheet_survey_75 extends cSheetITS
{
	private	$name = 'VM128';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Threats to Multi Residential Investment</u></strong><br />How serious are the following <strong><em>potential factors</em></strong> on the expected return on investment for Multi Residential investments in the future? (1 = least significant, 5 = most significant)</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'A', 'a) Unexpected Capital Costs', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B', 'b) Legislative Changes (rent control)', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'C', 'c) Debt Costs', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'D', 'd) Energy Costs', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'E', 'e) Realty Taxes', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'F', 'f) Economic Conditions', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'G', 'g) CMHC Insurance', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'H', 'h) Vacancy/Availability', '', cAcquirer::AcquireSignificance()));
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
