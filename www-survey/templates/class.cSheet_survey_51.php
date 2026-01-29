<?php
// BD - Feb 25, 2009

class cSheet_survey_51 extends cSheetITS
{
	private	$name = 'MI108';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the current year, how much does your company expect to:</p><p>a) <strong>Invest</strong> in additional acquisitions and or new development, or issue new debt in Canada?</p><p>b) <strong>Dispose</strong> of assets, or retire debt in Canada?</p></div>');

		$this->AddInputMatrix_PropertyType($this->name, 'Value');
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_PropertyType($this->name);
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
}

?>
