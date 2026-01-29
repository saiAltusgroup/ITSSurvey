<?php
// DK - May 4, 2009

class cSheet_survey_71 extends cSheetITS
{
	private	$name = 'VR27';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>G.R.O.C. Ratio</u></strong><br />What is an appropriate average target <strong>Gross Rent to Sales Ratio</strong> for CRU tenants within a <strong>Benchmark Tier I Regional Mall.</strong> Gross Rent comprises base rent, C.A.M., taxes, HVAC, promotion, advertising, etc.</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name, 'Gross Rent to Sales Ratio', '', cAcquirer::AcquireGROCRatio()));
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
