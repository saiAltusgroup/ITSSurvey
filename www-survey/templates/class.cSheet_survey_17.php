<?php
// DK - Oct 22, 2008

class cSheet_survey_17 extends cSheetITS
{
	private	$name = 'VR124';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Tenant Retention Ratio</u></strong><br />What tenant retention ratio (CRU % renew only) is utilized for the following retail property types?</p></div>');

//		$this->AddInput(new cInputTextBox('s17_124aVPR', 'Tenant Retention Ratio', ''));

		$this->AddInputMatrix_tenantRetention($this->name, 'Value');
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'A');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'B');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'E');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_tenantRetention($this->name);
	}
}

?>
