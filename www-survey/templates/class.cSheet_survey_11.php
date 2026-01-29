<?php
// DK - Oct 22, 2008

class cSheet_survey_11 extends cSheetITS
{
	private	$name = 'VO132';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>In establishing a net effective rental rate, do you discount the cash flows <strong>monthly</strong> or <strong>annually</strong>? What is the discount rate you use to amortize tenant concessions, free rent or other form of inducements?</p></div>');

		$this->AddInput(new cInputChoiceSingleRadiolist($this->name . 'B', 'Discount Period', '', cAcquirer::AcquireDiscountPeriod()));
		$this->AddInput(new cInputTextBox($this->name . 'A', 'Discount Rate (%)', ''));

	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$validator = new cValidator(true);
		$input=$this->FindInput($this->name . 'A');
		if (strlen(trim($input->GetValue()))>0) $ret &= $validator->ValidateNumericRange($input, 1, 100, _ERROR_MUST_BE_RATE);
		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
		parent::ParentDisplay();
	}
}

?>
