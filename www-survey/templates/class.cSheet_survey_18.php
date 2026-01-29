<?php
// DK - Oct 22, 2008

class cSheet_survey_18 extends cSheetITS
{
	private	$name = 'VR59';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Retail Lag Vacancy</u></strong><br />What lag vacancy (including fixturing period), expressed in number of months, is utilized for new non-anchored tenants in the following retail property types?</p></div>');

//		$this->AddInput(new cInputTextBox('s18_59aVPR', 'Retail Lag Vacancy', ''));

		$this->AddInputMatrix_LagVacancy($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 1, 1000000, _ERROR_MUST_BE_IN_NUMERIC_RANGE2);
			if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateInteger($input, _ERROR_MUST_BE_INTEGER);
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_LagVacancy($this->name);
	}
}

?>
