<?php
// BD - Feb 24, 2009

class cSheet_survey_50 extends cSheetITS
{
	private	$name = 'CD85A';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>Will the cost of debt for conventional first mortgages (i.e. the combination of a. - general interest rates and b. - rate spreads over 10 year Canadian Government Bonds), increase, decrease or stay the same in the next three months for each of the property types below? (Pick one number below)</p></div>');

		$this->AddInputMatrix_arrDebtCost($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_arrDebtCost($this->name, 'Debt Cost Outlook – Conventional First Mortgage');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();
		$validator = new cValidator(true);

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, '');
		}
		return $this->Validate_Ignored($ret);
		
	}
}

?>
