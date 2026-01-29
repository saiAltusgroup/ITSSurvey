<?php
// DK - May 4, 2009

class cSheet_survey_59 extends cSheetITS
{
	private	$name = 'MI49';

	public function /* void */ Create()
	{

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . '', 'Do you deduct selling costs in calculation of reversion value?', '', cAcquirer::AcquireYesNoSometimes()));
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
