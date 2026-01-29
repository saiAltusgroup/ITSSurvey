<?php
// DK - Apr 9, 2010 - version 1.0

require_once(_FORMAPI_PATH . 'classes/class.cDriverExportEmail.php');

class cDriverExportEmailInsite extends cDriverExportEmail
{
	private $Username = 'license@altusgroup.com';
	private $APIKey = '696fff36-2b22-4fd5-8d96-f1b0fe19aa13';

public function __construct($from, $to,$headersub, $replyTo, $subject, $bodyText, $password = null, $username = null)
    {
        $this->MailFrom = $from;
        $this->MailTo = $to;
        $this->ReplyTo = $replyTo;
$this->MailSubject = $headersub;
        $this->MailSubject = $subject;
        $this->MailBody = $bodyText;
       
        $this->Password = $password;
        $this->Username = $username;
    }
	public function Mail($filepath = null, $filename = null)
	{
		$attachment_id = '';
		if (isset($filename)) $attachment_id = $this->UploadAttachment($filepath, $filename);

		$res = "";

		$data = "username=".urlencode($this->Username);
		$data .= "&api_key=".urlencode($this->APIKey);
		$data .= "&from=".urlencode($this->MailFrom);
		$data .= "&from_name=".urlencode($this->MailFrom);
		$data .= "&to=".urlencode($this->MailTo);
		$data .= "&subject=".urlencode($this->MailSubject);
		if ($this->MailBody) $data .= "&body_text=" . urlencode($this->MailBody);
		if (strlen($attachment_id) > 0) $data .= "&attachments=" . urlencode($attachment_id);

		$header = "POST /mailer/send HTTP/1.0\r\n";
		$header .= "Content-Type: application/x-www-form-urlencoded\r\n";
		$header .= "Content-Length: " . strlen($data) . "\r\n\r\n";
		$fp = fsockopen('ssl://api.elasticemail.com', 443, $errno, $errstr, 30);

		if (!$fp) return "ERROR. Could not open connection";
		else
		{
			fputs ($fp, $header . $data);
			while (!feof($fp))
			{
				$res .= fread($fp, 1024);
			}
			fclose($fp);
		}
		return $res;
	}

	public function UploadAttachment($filepath, $filename)
	{
		$res = "";
		$data = file_get_contents($filepath);

		$header = "PUT /attachments/upload?username=" . $this->Username . "&api_key=" . $this->APIKey . "&file=" . $filename . " HTTP/1.0\r\n";
		$header .= "Content-Type: application/x-www-form-urlencoded\r\n";
		$header .= "Content-Length: " . strlen($data) . "\r\n\r\n";
		$fp = fsockopen("ssl://api.elasticemail.com", 443, $errno, $errstr, 30);

		if (!$fp) return "ERROR. Could not open connection";
		else
		{
			fputs($fp, $header.$data);
			while (!feof($fp))
			{
				$res .= fread ($fp, 1024);
			}
			fclose($fp);
		}

		$res = substr($res, -9);

		return $res;
	}
}

?>
