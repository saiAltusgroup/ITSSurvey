<?php
// DK - Oct 22, 2008

class cSheet_survey_26 extends cSheetITS
{
	private	$name = 'VM82';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>For the <strong>Benchmark Suburban Multi-Unit Residential</strong> rental building, what is the cost of preparing a typical two-bedroom suite for tenant move-in. The cost is NOT for a major capital upgrade, but simply the cost of bringing the suite to a leaseable condition (e.g. cleaning, painting, carpeting, appliances and minor repairs).</p></div>');

//		$this->AddInput(new cInputTextBox('s26_82VPMR', 'Cost $ per Suite', ''));

		$this->AddInputMatrix_SuburbanMultiSuiteCost($this->name, 'Value');

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

	public function /* void */ CustomDisplay()
	{
//		parent::ParentDisplay();

		$this->DrawInputMatrix_SuburbanMultiSuiteCost($this->name);
	}
}

?>
