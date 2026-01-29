<?php
// DK - May 4, 2009

class cSheet_survey_61 extends cSheetITS
{
	private	$name = 'MI89';

	public function /* void */ Create()
	{
		$this->SetHeader('<p><strong>Financial Indicators</strong></p><div class="descr"><p>For the benchmark Downtown Class "AA" Office Building, rank the following financial indicators, 1 through 6 (with 1 being the most important):<br /><strong><u>No Duplication</u></strong></p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'A', 'a. DCF / IRR', '', cAcquirer::AcquireRank()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B', 'b. NOI Yield Year 1', '', cAcquirer::AcquireRank()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'C', 'c. NOI Yield Average Years 1 - 5', '', cAcquirer::AcquireRank()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'D', 'd. CF Yield Year 1', '', cAcquirer::AcquireRank()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'E', 'e. CF Yield Average Years 1 - 5', '', cAcquirer::AcquireRank()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'F', 'f. $ Per s.f.', '', cAcquirer::AcquireRank()));
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
