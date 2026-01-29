<?php
// DK - Oct 22, 2008

class cSheet_survey_12 extends cSheetITS
{
	private	$name = 'VO103';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>Over the next three months, will vacancy increase, decrease, or stay the same for the following products / markets. Please insert code 1 = increase, code 2 = decrease, code 3 = no change.</p></div>');

		$this->AddInputMatrix_VacancyBarometer($this->name, 'Value');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, '');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_VacancyBarometer($this->name);
	}
}

?>
