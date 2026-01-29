<?php
// DK - May 4, 2009

class cSheet_survey_63 extends cSheetITS
{
	private	$name = 'MI18';


	public function /* void */ Create()
	{
		$this->SetHeader('<p><strong>Interest Rate Impact on Capitalization Rates</strong></p><div class="descr"><p>Given the current investment environment, what impact would a 1.0% increase in interest rates have on overall capitalization rates (assuming same loan-to-value ratio)?</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name, 'Interest Rate Impact on Capitalization Rates (basis points):', '', cAcquirer::AcquireImpact()));
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
