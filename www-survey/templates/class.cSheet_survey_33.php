<?php
// BD - Feb 23, 2009

class cSheet_survey_33 extends cSheetITS
{
	private	$name = 'VO130';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For a <strong><em>newly constructed</em> Downtown Class "AA" and Suburban Class "A" Office Building in <em>Toronto</em></strong>, what are the current "<strong>MARKET</strong>" valuation parameters and Value Trend for the next 12 months?</p><p>The Downtown Class "AA" and Suburban Class "A" building is defined as:</p><ul><li>	Downtown "AA" building is approximately 600,000 s.f. with LEED Silver designation</li><li>	Suburban "A" building is approximately 100,000 s.f. with LEED Silver designation</li><li>	100% leased to a diversified mix of investment grade tenancies at market rents with contractual rental escalations every 5 years.</li><li>	10 year lease term</li><li>	Downtown "AA" rental rates: $30.00 psf (yrs.1-5 with an escalation to $34.00 psf (years 6-10)., based on an increase of 2.5% per annum compounded.</li><li>	Suburban "A" rental rates: $20.00 psf (yrs.1-5) with an escalation to $ 22.75 psf (years 6-10).</li></ul></div>');
				 
		$this->AddInputMatrix_arrNewConClass($this->name, 'Value');
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
			$ret &= $this->Validate_Rate($input, _ERROR_MUST_BE_RATE, 'G_');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'H_');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
		//parent::ParentDisplay();

		$this->DrawInputMatrix_ArrNewConst($this->name);
	}
}

?>
