<?php
// DK - Oct 22, 2008

class cAcquirer
{
	public function /* cChoicelist */ AcquireYesNo()
	{
		$choice_array = array('yes' => _YES, 'no' => _NO);
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireYesNoSometimes()
	{
		$choice_array = array('1' => 'Yes', '2' => 'No', '3' => 'Sometimes');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireMethod()
	{
		$choice_array = array('1' => 'Method 1', '2' => 'Method 2', '3' => 'Method 3');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireImpact()
	{
		$choice_array = array('1' => 'No Impact', '2' => 'Increase of 25 bps', '3' => 'Increase of 50 bps', '4' => 'Increase of 75 bps', '5' => 'Increase of 100 bps', '6' => 'More than 100 bps');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireDiscountPeriod()
	{
		$choice_array = array('1' => 'Annual', '2' => 'Monthly', '9' => 'Don\'t Know / Not Sure');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireSignificance()
	{
		$choice_array = array('1' => '1 = Least Significant', '2' => '2', '3' => '3', '4' => '4', '5' => '5 = Most Significant');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireRating()
	{
		$choice_array = array('1' => '1 = Very Poor', '2' => '2 = Poor', '3' => '3 = Good', '4' => '4 = Very Good', '5' => '5 = Excellent');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireRank()
	{
		$choice_array = array('1' => '1 = most important', '2' => '2', '3' => '3', '4' => '4', '5' => '5', '6' => '6 = least important');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireIndustry()
	{
		$choice_array = array('1' => 'Ownership Side of Pension Funds / Pension Fund Advisors', '2' => 'Publicly Traded Corporations / REITs', '3' => 'Ownership Side of Banks / Life Companies', '4' => 'Private Companies - Domestic', '5' => 'Private Companies - Foreign (US, Europe, Asia etc)', '6' => 'Investment Banks / Real Estate Brokers / Real Estate Analysts', '7' => 'Lending Side of Banks / Life Companies / Pension Funds', '8' => 'Other');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireExpertise()
	{
		//$choice_array = array('1' => 'Office', '2' => 'Retail', '3' => 'Industrial', '4' => 'Multi-Residential', '5' => 'Capital Sources', '6' => 'Miscellaneous');
		$choice_array = array('1' => 'Office', '2' => 'Retail', '3' => 'Industrial', '4' => 'Multi-Residential', '5' => 'Capital Sources');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireMarkets()
	{
		$choice_array = array('1' => 'Vancouver', '2' => 'Edmonton', '3' => 'Calgary', '4' => 'Toronto', '5' => 'Ottawa', '6' => 'Montreal', '7' => 'Quebec City', '8' => 'Halifax');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireRetailMgmtFees()
	{
		$choice_array = array('1' => '<2.00%', '2' => '2.01% - 2.50%', '3' => '2.51% - 3.00%', '4' => '3.01% - 3.50%', '5' => '>3.51%');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireGROCRatio()
	{
		$choice_array = array('1' => '10.00% to 11.00%', '2' => '11.01% to 12.00%', '3' => '12.01% to 13.00%', '4' => '13.01% to 14.00%', '5' => '14.01% to 15.00%', '6' => '15.01% to 16.00%', '7' => '16.01% to 17.00%', '8' => '17.01% to 18.00%', '9' => '>18.01');
		return new cChoicelist($choice_array);
	}

	public function /* cChoicelist */ AcquireMortgageCompanies()
	{
		$choice_array = array('0' => 'Please select', '1' => 'Bank of Montreal', '2' => 'Bank of Nova Scotia', '3' => 'CIBC', '4' => 'GWL', '5' => 'G.E. Capital', '6' => 'GMAC', '7' => 'Manulife', '8' => 'Mass Mutual', '9' => 'Metropolitan Life', '10' => 'Merrill Lynch', '11' => 'Sun Life', '12' => 'OMERS', '13' => 'Royal Bank', '14' => 'Standard Life', '15' => 'TIAA - CREF', '16' => 'Toronto Dominion Bank / CMO', '17' => 'Caisse de Depot / CDPQ / MCAP', '18' => 'Others - Must Specify');
		return new cChoicelist($choice_array);
	}
}
?>
