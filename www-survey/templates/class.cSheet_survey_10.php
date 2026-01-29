<?php
// DK - Oct 22, 2008

class cSheet_survey_10 extends cSheetITS
{
	private	$name = 'VO128';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong>Only</strong> answer the following question if you:</p><ul><li>provide third party <strong><u>property management</u></strong> services, or</li><li>outsource third party <strong><u>property management</u></strong> services</li></ul><p>What is an appropriate <strong>third party property management fee</strong>, expressed as percentage (i.e. 0.5, 1.5, 1.75 etc.) for an office property, excluding leasing fees and commissions?</p></div>');

//		$this->AddInput(new cInputTextBox('s10_128VPO_A', '< 5', ''));
//		$this->AddInput(new cInputTextBox('s10_128VPO_B', '5 – 10', ''));
//		$this->AddInput(new cInputTextBox('s10_128VPO_C', '10 – 20', ''));
//		$this->AddInput(new cInputTextBox('s10_128VPO_D', '20 – 50', ''));
//		$this->AddInput(new cInputTextBox('s10_128VPO_E', '50 – 100', ''));
//		$this->AddInput(new cInputTextBox('s10_128VPO_F', '> 100', ''));

		$this->AddInputMatrix_Arr9($this->name, 'Value');

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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'F');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_Arr9($this->name);
	}
}

?>
