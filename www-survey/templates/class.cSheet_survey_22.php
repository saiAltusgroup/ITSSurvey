<?php
// DK - Oct 22, 2008

class cSheet_survey_22 extends cSheetITS
{
	private	$name = 'VI133';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Net Face Rental Rate</u></strong><br />For a typical lease in a <strong>Benchmark Multi-Tenant Industrial Building</strong>, what would be the typical year 1 - 5 net face rental rate per square foot?</p></div>');

//		$this->AddInput(new cInputTextBox('s22_133VPI', '$ Psf', ''));
		$this->AddInputMatrix_InitialNetFace($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 0, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE);
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_InitialNetFace($this->name);
	}
}

?>
