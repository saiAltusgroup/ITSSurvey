<?php
// DK - May 4, 2009

class cSheet_survey_69 extends cSheetITS
{
	private	$name = 'VO131';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Motivation for Office Investment</u></strong><br />At current and expected pricing for investment in this sector (stabilized building) what is the strength of the following factors underpinning the motivation for an acquisition? (1 = least significant, 5 = most significant)</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'A', 'a) Initial Return', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B', 'b) Cash Flow Upside (Growth)', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'C', 'c) Capital Appreciation', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'D', 'd) Requirement to Place Capital', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'E', 'e) Poor Performance of Alternative Investments', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'F', 'f) Availability and Low Cost of Debt/Equity', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'G', 'g) Enhances/Upgrades Existing Portfolio', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'H', 'h) Operating Management/Efficiencies (i.e. Economies of Scale', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'I', 'i) Flight Capital', '', cAcquirer::AcquireSignificance()));
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
