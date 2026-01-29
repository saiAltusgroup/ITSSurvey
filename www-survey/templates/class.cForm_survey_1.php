<?php
// DK - Oct 22, 2008

//require_once('classes/class.ConverterQid.php');

class cForm_survey_1 extends cForm
{
	public function /* void */ Create()
	{
		$user = $GLOBALS['user'];
		$expertise_arr = split(',', $user->GetEXPERTISE());
		$markets_arr = split(',', $user->GetMARKETS());

		//echo 'Expertise: ' . $user->GetEXPERTISE() . '<br />';
		//echo 'Markets: ' . $user->GetMARKETS() . '<br />';

		if (_QTR_ID=='Q1')
		{
			/* I) Altus InSite Product / Market Barometer */
			$this->AddSheet(new cSheet_survey_1('sheet_1', '1 - 97APMB'));

			/* II) Valuation Parameters - Office */
			if (in_array(1, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_2('sheet_2', '2 - 99VPO'));
				$this->AddSheet(new cSheet_survey_32('sheet_3', '3 - 55VPO'));
				$this->AddSheet(new cSheet_survey_33('sheet_4', '4 - 130VPO'));
				$this->AddSheet(new cSheet_survey_34('sheet_5', '5 - 127VPO'));
				$this->AddSheet(new cSheet_survey_35('sheet_6', '6 - 122VPO'));
				$this->AddSheet(new cSheet_survey_36('sheet_7', '7 - 123VPO'));
				if (in_array(6, $markets_arr))
				{
					$this->AddSheet(new cSheet_survey_7('sheet_8', '8 - 119VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_8('sheet_9', '9 - 120VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_9('sheet_10', '10 - 121VPO')); /* Montreal only question */
				}
				$this->AddSheet(new cSheet_survey_40('sheet_11', '11 - 3VPO'));
				$this->AddSheet(new cSheet_survey_12('sheet_12', '12 - 103VPO'));
			}

			/* III) Valuation Parameters - Retail */
			if (in_array(2, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_13('sheet_13', '13 - 100VPR'));
				$this->AddSheet(new cSheet_survey_37('sheet_14', '14 - 61VPR'));
				$this->AddSheet(new cSheet_survey_38('sheet_15', '15 - 69VPR'));
				$this->AddSheet(new cSheet_survey_39('sheet_16', '16 - 125VPR'));
				$this->AddSheet(new cSheet_survey_41('sheet_17', '17 - 70VPR'));
				$this->AddSheet(new cSheet_survey_42('sheet_18', '18 - 127VPR'));
				$this->AddSheet(new cSheet_survey_43('sheet_19', '19 - 34VPR'));
			}

			/* IV) Valuation Parameters - Industrial */
			if (in_array(3, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_20('sheet_20', '20 - 75VPI'));
				$this->AddSheet(new cSheet_survey_21('sheet_21', '21 - 33VPI'));
				$this->AddSheet(new cSheet_survey_44('sheet_22', '22 - 128VPI'));
				$this->AddSheet(new cSheet_survey_45('sheet_23', '23 - 129VPI'));
			}

			/* V) Valuation Parameters - Multi-Residential */
			if (in_array(4, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_24('sheet_24', '24 - 36VPMR'));
				$this->AddSheet(new cSheet_survey_46('sheet_25', '25 - 38VPMR'));
				$this->AddSheet(new cSheet_survey_47('sheet_26', '26 - 117VPMR'));
				$this->AddSheet(new cSheet_survey_48('sheet_27', '27 - 127VPMR'));
			}

			/* VI) Capital Sources - Equity Market Focus */
			if (in_array(5, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_49('sheet_28', '28 - 83CSDM'));
				$this->AddSheet(new cSheet_survey_50('sheet_29', '29 - 85aCSDM'));
			}

			/* VII) Miscellaneous */
			//if (in_array(6, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_51('sheet_30', '30 - 108MISC'));
				$this->AddSheet(new cSheet_survey_31('sheet_31', '31 - 93MISC'));
			}
		}

		else if (_QTR_ID=='Q2')
		{
			/* I) Altus InSite Product / Market Barometer */
			$this->AddSheet(new cSheet_survey_1('sheet_1', '1 - 97APMB'));

			/* II) Valuation Parameters - Office */
			if (in_array(1, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_2('sheet_2', '2 - 99VPO'));
				$this->AddSheet(new cSheet_survey_3('sheet_3', '3 - 21VPO'));
				$this->AddSheet(new cSheet_survey_4('sheet_4', '4 - 22VPO'));
				$this->AddSheet(new cSheet_survey_52('sheet_5', '5 - 136VPO'));
				$this->AddSheet(new cSheet_survey_53('sheet_6', '6 - 23VPO'));
				$this->AddSheet(new cSheet_survey_54('sheet_7', '7 - 134VPO'));
				$this->AddSheet(new cSheet_survey_55('sheet_8', '8 - 135VPO'));
				$this->AddSheet(new cSheet_survey_6('sheet_9', '9 - 25VPO'));

				if (in_array(6, $markets_arr))
				{
					$this->AddSheet(new cSheet_survey_7('sheet_10', '10 - 119VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_8('sheet_11', '11 - 120VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_9('sheet_12', '12 - 121VPO')); /* Montreal only question */
				}

				$this->AddSheet(new cSheet_survey_12('sheet_13', '13 - 103VPO'));
			}

			/* III) Valuation Parameters - Retail */
			if (in_array(2, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_13('sheet_14', '14 - 100VPR'));
				$this->AddSheet(new cSheet_survey_14('sheet_15', '15 - 101VPR'));
				$this->AddSheet(new cSheet_survey_16('sheet_16', '16 - 28VPR'));
				$this->AddSheet(new cSheet_survey_56('sheet_17', '17 - 126VPR'));
				$this->AddSheet(new cSheet_survey_57('sheet_18', '18 - 30VPR'));
			}

			/* IV) Valuation Parameters - Industrial */
			if (in_array(3, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_20('sheet_19', '19 - 75VPI'));
				$this->AddSheet(new cSheet_survey_21('sheet_20', '20 - 33VPI'));
				$this->AddSheet(new cSheet_survey_22('sheet_21', '21 - 133VPI'));
				$this->AddSheet(new cSheet_survey_23('sheet_22', '22 - 134VPI'));
			}

			/* V) Valuation Parameters - Multi-Residential */
			if (in_array(4, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_24('sheet_23', '23 - 36VPMR'));
				$this->AddSheet(new cSheet_survey_58('sheet_24', '24 - 40VPMR'));
			}

			/* VI) Capital Sources - Equity Market Focus */
			if (in_array(5, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_28('sheet_25', '25 - 47CSCM'));
				$this->AddSheet(new cSheet_survey_29('sheet_26', '26 - 48CSCM'));
				$this->AddSheet(new cSheet_survey_30('sheet_27', '27 - 43CSCM'));
			}

			/* VII) Miscellaneous */
			//if (in_array(6, $expertise_arr))
			{
				/* Note for Q2 2010: Please discontinue asking the question and producing the chart for 49MISC (reversion value) - requested by Kevin Antaya on Aug 18, 2009. */
				//$this->AddSheet(new cSheet_survey_59('sheet_28', '28 - 49MISC'));
				//$this->AddSheet(new cSheet_survey_60('sheet_29', '29 - 50MISC'));
				$this->AddSheet(new cSheet_survey_61('sheet_30', '28 - 89MISC'));
				//$this->AddSheet(new cSheet_survey_62('sheet_31', '31 - 11MISC'));
				$this->AddSheet(new cSheet_survey_63('sheet_32', '29 - 18MISC'));
				$this->AddSheet(new cSheet_survey_64('sheet_33', '30 - 110MISC'));
				$this->AddSheet(new cSheet_survey_31('sheet_34', '31 - 93MISC'));
			}
		}

		else if (_QTR_ID=='Q3')
		{
			/* I) Altus InSite Product / Market Barometer */
			$this->AddSheet(new cSheet_survey_1('sheet_1', '1 - 97APMB'));

			/* II) Valuation Parameters - Office */
			if (in_array(1, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_2('sheet_2', '2 - 99VPO'));
				$this->AddSheet(new cSheet_survey_65('sheet_3', '3 - 64VPO'));
				$this->AddSheet(new cSheet_survey_66('sheet_4', '4 - 65VPO'));
				$this->AddSheet(new cSheet_survey_67('sheet_5', '5 - 66VPO'));
				$this->AddSheet(new cSheet_survey_68('sheet_6', '6 - 67VPO'));
				$this->AddSheet(new cSheet_survey_33('sheet_7', '7 - 130VPO'));
				if (in_array(6, $markets_arr))
				{
					$this->AddSheet(new cSheet_survey_7('sheet_8', '8 - 119VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_8('sheet_9', '9 - 120VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_9('sheet_10', '10 - 121VPO')); /* Montreal only question */
				}
				$this->AddSheet(new cSheet_survey_69('sheet_11', '11 - 131VPO'));
				$this->AddSheet(new cSheet_survey_70('sheet_12', '12 - 4VPO'));
				$this->AddSheet(new cSheet_survey_12('sheet_13', '13 - 103VPO'));
			}

			/* III) Valuation Parameters - Retail */
			if (in_array(2, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_13('sheet_14', '14 - 100VPR'));
				$this->AddSheet(new cSheet_survey_71('sheet_15', '15 - 27VPR'));
				$this->AddSheet(new cSheet_survey_38('sheet_16', '16 - 69VPR'));
				$this->AddSheet(new cSheet_survey_72('sheet_17', '17 - 139VPR'));
				$this->AddSheet(new cSheet_survey_41('sheet_18', '18 - 70VPR'));
				$this->AddSheet(new cSheet_survey_73('sheet_19', '19 - 33VPR'));
			}

			/* IV) Valuation Parameters - Industrial */
			if (in_array(3, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_20('sheet_20', '20 - 75VPI'));
				$this->AddSheet(new cSheet_survey_44('sheet_21', '21 - 128VPI'));
				$this->AddSheet(new cSheet_survey_45('sheet_22', '22 - 129VPI'));
				$this->AddSheet(new cSheet_survey_21('sheet_23', '23 - 33VPI'));
			}

			/* V) Valuation Parameters - Multi-Residential */
			if (in_array(4, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_24('sheet_24', '24 - 36VPMR'));
				$this->AddSheet(new cSheet_survey_74('sheet_25', '25 - 81VPMR'));
				$this->AddSheet(new cSheet_survey_48('sheet_26', '26 - 127VPMR'));
				$this->AddSheet(new cSheet_survey_75('sheet_27', '27 - 128VPMR'));
			}

			/* VI) Capital Sources - Equity Market Focus */
			if (in_array(5, $expertise_arr))
			{
				//$this->AddSheet(new cSheet_survey_76('sheet_28', '28 - 82CSDM'));
				$this->AddSheet(new cSheet_survey_49('sheet_29', '28 - 83CSDM'));
				$this->AddSheet(new cSheet_survey_50('sheet_30', '29 - 85aCSDM'));
			}

			/* VII) Miscellaneous */
			//if (in_array(6, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_77('sheet_31', '30 - 96MISC'));
				$this->AddSheet(new cSheet_survey_31('sheet_32', '31 - 93MISC'));
			}
		}
		else if (_QTR_ID=='Q4')
		{
			/* I) Altus InSite Product / Market Barometer */
			$this->AddSheet(new cSheet_survey_1('sheet_1', '1 - 97APMB'));

			/* II) Valuation Parameters - Office */
			if (in_array(1, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_2('sheet_2', '2 - 99VPO'));
				$this->AddSheet(new cSheet_survey_3('sheet_3', '3 - 21VPO'));
				$this->AddSheet(new cSheet_survey_4('sheet_4', '4 - 22VPO'));
				$this->AddSheet(new cSheet_survey_5('sheet_5', '5 - 57VPO'));
				$this->AddSheet(new cSheet_survey_6('sheet_6', '6 - 25VPO'));

				if (in_array(6, $markets_arr))
				{
					$this->AddSheet(new cSheet_survey_7('sheet_7', '7 - 119VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_8('sheet_8', '8 - 120VPO')); /* Montreal only question */
					$this->AddSheet(new cSheet_survey_9('sheet_9', '9 - 121VPO')); /* Montreal only question */
				}

				$this->AddSheet(new cSheet_survey_10('sheet_10', '10 - 128VPO'));
				$this->AddSheet(new cSheet_survey_11('sheet_11', '11 - 132VPO'));
				$this->AddSheet(new cSheet_survey_12('sheet_12', '12 - 103VPO'));
			}

			/* III) Valuation Parameters - Retail */
			if (in_array(2, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_13('sheet_13', '13 - 100VPR'));
				$this->AddSheet(new cSheet_survey_14('sheet_14', '14 - 101VPR'));
				$this->AddSheet(new cSheet_survey_15('sheet_15', '15 - 102VPR'));
				$this->AddSheet(new cSheet_survey_16('sheet_16', '16 - 28VPR'));
				$this->AddSheet(new cSheet_survey_17('sheet_17', '17 - 124aVPR'));
				$this->AddSheet(new cSheet_survey_18('sheet_18', '18 - 59aVPR'));
				$this->AddSheet(new cSheet_survey_19('sheet_19', '19 - 60VPR'));
			}

			/* IV) Valuation Parameters - Industrial */
			if (in_array(3, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_20('sheet_20', '20 - 75VPI'));
				$this->AddSheet(new cSheet_survey_21('sheet_21', '21 - 33VPI'));
				$this->AddSheet(new cSheet_survey_22('sheet_22', '22 - 133VPI'));
				$this->AddSheet(new cSheet_survey_23('sheet_23', '23 - 134VPI'));
			}

			/* V) Valuation Parameters - Multi-Residential */
			if (in_array(4, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_24('sheet_24', '24 - 36VPMR'));
				$this->AddSheet(new cSheet_survey_25('sheet_25', '25 - 118VPMR'));
				$this->AddSheet(new cSheet_survey_26('sheet_26', '26 - 82VPMR'));
				$this->AddSheet(new cSheet_survey_27('sheet_27', '27 - 129VPMR'));
			}

			/* VI) Capital Sources - Equity Market Focus */
			if (in_array(5, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_28('sheet_28', '28 - 47CSCM'));
				$this->AddSheet(new cSheet_survey_29('sheet_29', '29 - 48CSCM'));
				$this->AddSheet(new cSheet_survey_30('sheet_30', '30 - 43CSCM'));
			}

			/* VII) Miscellaneous */
			//if (in_array(6, $expertise_arr))
			{
				$this->AddSheet(new cSheet_survey_31('sheet_31', '31 - 93MISC'));
			}
		}

	}

	function IsCompleted()
	{
		// loops through all sheets in the form
		for ($i = 0; $i < $this->GetSheetCount(); $i++)
		{
			$sheet = $this->GetSheet($i);

			// loops through all inputs in the sheet
			for ($j = 0; $j < $sheet->GetInputCount(); $j++)
			{
				// checks if a non-Literal value exists and returns true if it finds one
				if ((strlen(trim($sheet->GetInput($j)->GetValue())) > 0) && ($sheet->GetInput($j)->GetUniqueID() != 'InputLiteral'))
				{
					return true;
				}
			}
		}

		return false;
	}
}
?>