<?php
// DK - Oct 22, 2008

class cSheet_survey_1 extends cSheetITS
{
	private	$name = 'Q1';

	public function /* void */ Create()
	{
		$this->SetHeader('<div class="descr"><p>We have developed this question to help establish changes in momentum for various products or markets. If you <em><strong>had</strong></em> to be a buyer (bullish) or seller (bearish), at current market pricing (holding is not an option), what would you do? (Insert Codes, \'buy = code 1\', \'sell = code 2\' or leave blank for No Comment or Opinion).</p></div>');

		//$this->AddInput(new cInputTextBox($this->name, 'Code', ''));

		$this->AddInputMatrix_Property($this->name, 'Value');

		$this->AddInput(new cInputTextBox($this->name . 'Q', 'REIT Units', ''));
		$this->FindInput($this->name . 'Q')->SetSize(1);
		$this->FindInput($this->name . 'Q')->SetMaxlength(1);

		$this->SetFooter('<div class="descr"><ul><li>Buy = Equates to a net buy, i.e. acquisitions are greater than dispositions</li><li>Sell = Equates to a net sell, i.e. dispositions are greater than acquisitions</li><li>No Comment / Not Familiar means you have no firm opinion</li></ul></div>');
	}

	public function /* cBool */ Validate()
	{
		$ret = true;
		$n = $this->GetInputCount();

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);

			//$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_9_OR_EMPTY, 9, '');
			$ret &= $this->Validate_1_2_x_or_empty($input, _ERROR_MUST_BE_1_2_OR_EMPTY, 1, '');
		}

		return $this->Validate_Ignored($ret);
	}

	public function /* void */ CustomDisplay()
	{
		//parent::ParentDisplay();

		$this->DrawInputMatrix_Property($this->name);
	}
}

?>
