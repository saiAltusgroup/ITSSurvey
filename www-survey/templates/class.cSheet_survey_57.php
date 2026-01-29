<?php
// DK - May 4, 2009

class cSheet_survey_57 extends cSheetITS
{
	private	$name = 'VR30';

	public function /* void */ Create()
	{
		$this->SetHeader('<p><strong>Retail Management Fees - % of Effective Gross Income</strong></p><div class="descr"><p><strong><u>Management Fees</u></strong><br />What is an appropriate third party management fee expressed as a percentage of effective gross income, (not including leasing fees and commissions) for the following benchmark retail property types:</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'A', 'Management Fees: Tier I Regional Mall', '', cAcquirer::AcquireRetailMgmtFees()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B', 'Management Fees: Enclosed Community Mall', '', cAcquirer::AcquireRetailMgmtFees()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'C', 'Management Fees: Power Centre', '', cAcquirer::AcquireRetailMgmtFees()));
		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'D', 'Management Fees: Food Anchored Retail Strip', '', cAcquirer::AcquireRetailMgmtFees()));
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
