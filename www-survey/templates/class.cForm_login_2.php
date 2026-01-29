<?php
// DK - Oct 22, 2008

class cForm_login_2 extends cForm
{
	public function /* void */ Create()
	{
		$this->AddSheet(new cSheet_login_2('sheet_2', _SHEET_LOGIN));
	}
}
?>
