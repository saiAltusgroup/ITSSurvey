<?php
// DK - Oct 22, 2008

class cSheet_survey_31 extends cSheetITS
{
	private	$name = 'MISC93';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>Do you have any specific topics you would like us to address in the next survey?</p><p><em>The source of the questions will remain strictly confidential.</em></p><p>Would you have any special comments or remarks concerning the current real estate market (general, location or property-type specific)?</p><p><em>Frequent remarks may be edited and published for sharing amongst survey subscribers but the name of the respondents will be kept confidential.</em></p></div>');

		$this->AddInput(new cInputTextArea($this->name, 'Comments / Remarks', ''));
	}

	public function /* cBool */ Validate()
	{
		return true;
		//return $this->Validate_Ignored(true);
	}

	public function /* void */ CustomDisplay()
	{
		parent::ParentDisplay();
	}
}

?>
