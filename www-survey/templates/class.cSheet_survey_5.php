<?php
// DK - Oct 22, 2008

class cSheet_survey_5 extends cSheetITS
{
	private	$name = 'VO57';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>What is the economic face rate and development yield for a Class "AA" downtown building?</p><p><strong><em>(Development yield is defined as face rate total costs including hard and soft costs and land). Please use decimals where possible.</em></strong></p></div>');

//		$this->AddInput(new cInputTextBox('s5_57VPO_A', 'a. Economic Face Rate', ''));
//		$this->AddInput(new cInputTextBox('s5_57VPO_B', 'b. Development Yield', ''));

		$this->AddInputMatrix_Arr4($this->name, 'Value');

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
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
		//parent::ParentDisplay();

		$this->DrawInputMatrix_Arr4($this->name);
	}
}

?>
