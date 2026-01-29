<?php
// BD - Feb 23, 2009

class cSheet_survey_40 extends cSheetITS
{
	private	$name = 'VO3';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What tenant retention ratio (% renewal only) is utilized?</p></div>');

		$this->AddInputMatrix_VacancyBarometer_B($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
	$this->DrawInputMatrix_VacancyBarometer_B($this->name, 'Office Tenant Retention Ratio<br />% Renew Only');
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'A_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
