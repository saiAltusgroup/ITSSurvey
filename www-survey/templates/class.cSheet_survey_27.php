<?php
// DK - Oct 22, 2008

class cSheet_survey_27 extends cSheetITS
{
	private	$name = 'VM129';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Motivation for Multi Residential Investment</u></strong><br />At current and expected pricing for investment in this sector what is the strength of the following factors underpinning the motivating for an acquisition? (1 = least significant, 5 = most significant)</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'A', 'a) Debt Costs', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B', 'b) Potential Loan (debt) to Value', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'C', 'c) Stability of value versus other asset types', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'D', 'd) Capital Growth', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'E', 'e) Stability of income', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'F', 'f) Income growth', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'G', 'g) High Cost of new construction', '', cAcquirer::AcquireSignificance()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'H', 'h) CMHC Insurance', '', cAcquirer::AcquireSignificance()));
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
