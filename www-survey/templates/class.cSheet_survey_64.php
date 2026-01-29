<?php
// DK - May 4, 2009

class cSheet_survey_64 extends cSheetITS
{
	private	$name = 'MI110';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p><strong><u>Homogeneous Portfolio Effect</u></strong><br />Would a homogenous real estate portfolio located in the same general region attract a premium or a discount if sold on the open market today? Assume the following general features:</p><ul><li>Homogenous Portfolio of $150 Million acquired at the same time</li><li>Individual properties priced at market value</li><li>Reasonable quality investment grade properties (per Altus InSite benchmark definitions)</li><li>Similar region</li></ul><p>For each of the benchmark properties below, please select either a <strong>Discount (1)</strong> or <strong>Premium (2)</strong> and the corresponding percentage discount or premium.</p></div>');

		$this->AddInputMatrix($this->name, 'Value', $this->arrDiscountPremiumPercent, $this->prodtype);
	}

	public function /* void */ CustomDisplay()
	{
		$this->head_title='Homogeneous Portfolio Effect<br />(Discount or Premium)';
		$this->head_type='Product Type';
		$this->head_example='';
		$this->DrawInputMatrix($this->name, $this->arrDiscountPremiumPercent, $this->prodtype, null);
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
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_OR_EMPTY, 1, 'A_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
		}
		return $this->Validate_Ignored($ret);
	}
}

?>
