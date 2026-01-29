<?php
// BD - Feb 23, 2009

class cSheet_survey_34 extends cSheetITS
{
	private	$name = 'VO127';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Suburban Class "B" Office Building</strong>, please provide your opinion of current "<strong>MARKET</strong>" valuation parameters. Also provide your opinion on the Value Trend for this asset class for the next 12 months.</p><p>The suburban Class "B" building is defined as:</p><ul><li>100,000 - 170,000 s.f.</li><li>Class "B" physical features</li><li>3.0 stalls per 1,000 s.f. surface parking</li><li>Prominent position within established and well recognized suburban business park</li><li>Multi tenant</li><li>5% vacancy</li><li>10% roll-over per year</li><li>Leased at <strong>market rates</strong> to good covenant tenants</li></ul></div>');
		$this->AddInputMatrix_arrRateSF_1($this->name, 'Value');
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'C_');
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'D_');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'E_');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
		$this->DrawInputMatrix_Arr1_B($this->name, 'Suburban Class "B" Office Building', 2);
	}
}

?>
