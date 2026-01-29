<?php
// DK - May 4, 2009

class cSheet_survey_62 extends cSheetITS
{
	private	$name = 'MI11';

	public function /* void */ Create()
	{
		$this->SetHeader('<p><strong>Discounted Cash Flow</strong></p>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name, 'Within the Discounted Cash Flow, to what income do you apply the reversion (going-out) Terminal Capitalization Rate?', '', cAcquirer::AcquireMethod()));

		$this->SetFooter('<div class="descr"><ul><li><strong>Method 1</strong>: Capitalize NOI <strong>before</strong> structural allowance reserve, TIs (tenant improvements) and leasing commissions.</li><li><strong>Method 2</strong>: Capitalize the NOI <strong>after</strong> adjustments for structural allowance reserve but before TIs (tenant improvements) and leasing commissions.</li><li><strong>Method 3</strong>: Capitalize Cash flow after structural allowance reserve, TIs (tenant improvements) and leasing commissions.</li></ul></div>');
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
