<?php
// DK - Oct 22, 2008

class cSheet_survey_6 extends cSheetITS
{
	private	$name = 'VO25';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Downtown Land Parcel</strong>, please provide your opinion of <strong>"MARKET"</strong> valuation parameters in terms <strong><u>$ per s.f. of buildable</u></strong> area <strong>NOT</strong> $ per s.f. of land area.</p><p>The benchmark downtown land parcel is defined as:</p><ul><li>100% location in local market</li><li>One acre (209\' x 209\') (maybe smaller in some markets)</li><li>Has two major and one minor street frontage</li><li>Has zoning in place for 525,000 s.f. of commercial density or 12.0 times coverage</li><li>For the purpose of this survey, assume the buildings / locations mentioned in the following chart have the above mentioned characteristics</li></ul></div>');

//		$this->AddInput(new cInputTextBox('s6_25VPO_A', 'a. $ Per s.f. Buildable', ''));
//		$this->AddInput(new cInputTextBox('s6_25VPO_B', 'b. Value Trend next 12 months (Increase = 1, Decrease = 2, Same / No, Change = 3', ''));

		$this->AddInputMatrix_Arr5($this->name, 'Value');
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
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_3_OR_EMPTY, 3, 'B_');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
		//parent::ParentDisplay();

		$this->DrawInputMatrix_Arr5($this->name);
	}
}

?>
