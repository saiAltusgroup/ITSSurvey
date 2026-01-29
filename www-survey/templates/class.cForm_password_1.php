<?php
// DK - Oct 22, 2008

class cForm_password_1 extends cForm
{
	public function /* void */ Create()
	{
		$this->AddSheet(new cSheet_password_1('sheet_1', _SHEET_PASSWORD));
	}
}
?>
