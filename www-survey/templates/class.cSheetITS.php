<?php
// DK - Oct 22, 2008

class cSheetITS extends cSheet
{
	protected /* cString */ $head_title;
	protected /* cString */ $head_type;
	protected /* cString */ $head_example;

	/* --- COLUMN ARRAYS --- */

	protected $province =
	array(
		array('B', 'British Columbia'),
		array('A', 'Alberta'),
		array('K', 'Saskatchewan'),
		array('M', 'Manitoba'),
		array('O', 'Ontario'),
		array('Q', 'Quebec'),
		array('T', 'Atlantic Canada')
	);

	protected $province_example_1 =
	array(
		array('B', 'Orchard Park, Kelowna'),
		array('A', 'Medicine hat Mall, Medicine Hat'),
		array('S', 'Midtown Mall, Saskatoon'),
		array('M', 'Kildonan Place, Winnipeg'),
		array('O', 'Lambton Mall, Sarnia'),
		array('Q', 'Les Galeries Chagnon, Levis'),
		array('C', 'Regent Mall, Fredericton')
	);

	protected $prodtype =
	array(
		array('O', 'Office Class A Buildings'),
		array('B', 'Office Class B Buildings'),
		array('R', 'Retail Food Anchored or Power Centre'),
		array('I', 'Multi Tenant Industrial'),
		array('M', 'Multi Unit Residential Suburban')
	);

	protected $market =
	array(
		array('V', 'Vancouver'),
		array('E', 'Edmonton'),
		array('C', 'Calgary'),
		array('T', 'Toronto'),
		array('O', 'Ottawa'),
		array('M', 'Montreal'),
		array('Q', 'Quebec City'),
		array('H', 'Halifax')
	);

	protected $market_alt =
	array(
		array('VAN', 'Vancouver'),
		array('EDM', 'Edmonton'),
		array('CAL', 'Calgary'),
		array('TOR', 'Toronto'),
		array('OTT', 'Ottawa'),
		array('MTL', 'Montreal'),
		array('Q', 'Quebec City'),
		array('HAL', 'Halifax')
	);

	protected $market_example_1 =
	array(
		array('V', 'Park Place'),
		array('E', 'Manulife Centre'),
		array('C', 'Devon Tower'),
		array('T', 'Freehold Exchange Tower'),
		array('O', 'World Exchange Plaza'),
		array('M', '1981 McGill College Avenue'),
		array('Q', 'Edifice la Solidarit&eacute;'),
		array('H', 'Purdy\'s Wharf')
	);

	protected $market_example_2 =
	array(
		array('V', 'Burrard &amp; Dunsmuir'),
		array('E', 'Commerce Court Position'),
		array('C', 'Petroleum Club Block'),
		array('T', 'Unon Station or King / Bay'),
		array('O', 'Sparks / Bank'),
		array('M', 'McGill College / de Maisonneuve'),
		array('Q', 'Rene Levesque'),
		array('H', 'Triangle Site')
	);

	protected $market_example_3 =
	array(
		array('V', 'Oakridge Centre'),
		array('E', 'Southgate Mall'),
		array('C', 'Market Mall'),
		array('T', 'Sherway Gardens'),
		array('O', 'Bayshore Shopping Centre'),
		array('M', 'Carrefour Laval'),
		array('Q', 'Place Laurier'),
		array('H', 'Mic Mac Mall')
	);

	protected $market_example_4 =
	array(
		array('V', 'Semiahmoo Shopping Centre'),
		array('E', 'Capilano Mall'),
		array('C', 'Westbrook Mall'),
		array('T', 'Bridlewood Mall'),
		array('O', 'Billings Bridge'),
		array('M', 'Carrefour De La Pointe'),
		array('Q', 'Place Quatre-Bourgeois'),
		array('H', 'Bedford Place Mall') // used to be BM Property
	);

	protected $market_example_5 =
	array(
		array('V', 'Annacis Island'),
		array('E', 'Northwest'),
		array('C', 'South Airways'),
		array('T', 'Heartland'),
		array('O', 'Sheffield / Lancaster'),
		array('M', 'St. Laurent'),
		array('Q', 'Colbert Park'), // used to be Vanier Park
		array('H', 'Burnside Industrial Park')
	);

	protected $market_example_6 =
	array(
		array('V', 'Crestwood'),
		array('E', 'Northwest'),
		array('C', 'South Airways'),
		array('T', 'Western Business Park'),
		array('O', 'Sheffield / Lancaster'),
		array('M', 'St. Laurent'),
		array('Q', 'Metrobec'),
		array('H', '86 Troop Avenue') // used to be ING Troop Avenue
	);

	protected $market_example_7 =
	array(
		array('V', 'Knightsbridge'),
		array('E', 'Northwest'),
		array('C', 'South Airways'),
		array('T', 'Western Business Park'),
		array('O', 'Sheffield / Lancaster'),
		array('M', 'St. Laurent'),
		array('Q', 'Parc Metrobec'),
		array('H', 'Burnside Industrial Park')
	);


	protected $market_example_8 =
	array(
		array('V', 'Knightsbridge'),
		array('E', 'Northwest'),
		array('C', 'South Airways'),
		array('T', 'Western Business Park'),
		array('O', 'Sheffield / Lancaster'),
		array('M', 'St. Laurent'),
		array('Q', 'Parc Metrobec'),
		array('H', 'Burnside Industrial Park')
	);

	protected $market_example_9 =
	array(
		array('V', 'Panarama Court Apt. (Burnaby)'),
		array('E', 'West Edmonton Village'),
		array('C', 'Glenmore Trail &amp; Elbow Dr.'),
		array('T', 'Don Mills / Sheppard'),
		array('O', 'Lincoln Heights'),
		array('M', 'Lakeshore Tower, Ponte-Claire'),
		array('Q', 'La Volière'),
		array('H', 'Forest Green Apartments (Clayton Park)')
	);

	protected $market_example_10 =
	array(
		array('V', 'Lougheed Commerce Court'),
		array('E', 'Eastgate Business Park'),
		array('C', 'Macleod Trail &amp; Southland'),
		array('T', '140 Allstate Parkway'),
		array('O', 'Kanata North Park'),
		array('M', 'South Shore'),
		array('Q', 'Lebourgneuf '),
		array('H', 'Summit Park')
	);

	protected $market_example_11 =
	array(
		array('V', 'Langley Power'),
		array('E', 'Mayfield Common'),
		array('C', 'West Hills'),
		array('T', 'Heartland'),
		array('O', 'South Keys Shopping Centre'),
		array('M', 'Marche Central'),
		array('Q', 'Lebourgneuf '),
		array('H', 'Bayers Lake')
	);

	protected $market_example_12 =
	array(
		array('V', 'Kensington Square'),
		array('E', 'Market Place at Collingwood'),
		array('C', 'Dalhousie Station'),
		array('T', 'Kingsbury Centre'),
		array('O', 'Strandherd Crossing'),
		array('M', 'Brien Shopping Centre'),
		array('Q', 'Place L?Ormière'),
		array('H', 'Park West Centre')
	);

	protected $market_example_13 =
	array(
		array('V', 'Park Place'),
		array('E', 'Manulife Centre'),
		array('C', 'Devon Tower'),
		array('T', 'Exchange Tower'),
		array('O', 'World Exchange Plaza'),
		array('M', '1981 McGill College Avenue'),
		array('Q', 'Edifice la Solidarité'),
		array('H', 'Purdy\'s Wharf') // used to be Bedford Place Mall // used to be BM Property
	);

	protected $market_example_14 =
	array(
		array('V', '409 Granville Street'),
		array('E', 'Energy Square'),
		array('C', 'Canoxy Building'),
		array('T', '4 King Street West'),
		array('O', '255 Albert Street'),
		array('M', '2020 University Avenue'),
		array('Q', '1150 Claire Fontaine'),
		array('H', 'CIBC Building') // used to be Bedford Place Mall // used to be BM Property
	);

	protected $market_example_15 =
	array(
		array('V', 'Willingdon Park'),
		array('E', 'Eastgate Business Park'),
		array('C', 'MacLeod Trail &amp; Southland'),
		array('T', 'Valleywood Corporate Centre'),
		array('O', 'Kanata North Park'),
		array('M', 'South Shore'),
		array('Q', 'Lebourgneuf Boulevard'), // used to be Boulevard Laurier
		array('H', 'Summit Park')
	);

	protected $market_example_16 =
	array(
		array('V', 'East of Denman (West End)'),
		array('E', 'The David Thompson Apartments'),
		array('C', 'West Core'),
		array('T', 'Bay/Bloor'),
		array('O', 'Centre Town'),
		array('M', 'Le Cartier (Peel &amp; Sherbrooke)'),
		array('Q', '440 Père-Marquette (Les Habitats)'),
		array('H', 'Plaza 1881 (Brunswick St.)')
	);

	protected $benchmark =
	array(
		array('C', 'Benchmark Class "A" Office'),
		array('S', 'Benchmark Class "B" Office'),
		array('M', 'Benchmark Retail'),
		array('I', 'Benchmark Industrial'),
		array('R', 'Benchmark Multi-Unit Residential CMHC Insured'),
	);

	protected $arrPropertyTypeOD =
	array(
		array('1O', 'Acquisition &amp; Development ($ Mil)'),
		array('2O', 'Disposition ($ Mil)'),
		array('1D', 'Issue New Debt ($ Mil)'),
		array('2D', 'Retire Debt ($ Mil)')
	);

	protected $arrType =
	array(
		array('A', 'Downtown Class "AA" Office'),
		array('B', 'Tier I Regional Mall'),
		array('C', 'Multi Tenant Industrial'),
		array('D', 'Multi ?Unit Residential'),
	);

	protected $class_physical =
	array(
		array('D', 'Downtown Class "AA"'),
		array('S', 'Suburban Class "A"')
	);

	protected $pv_example_1 =
	array(
		//array('*', 'Place Vendome<br />5252 De Maisonneuve Blvd. West')
		array('*', '4100 Molson')
	);

	protected $pv_example_2 =
	array(
		array('*', 'Jardins Brittany - 25  Brittany Avenue')
	);

	protected $reits =
	array(
		array('A', 'REITs'),
		array('B', 'Publicly Traded Corporations')
	);

	protected $caps =
	array(
		array('A', 'Large Cap'),
		array('B', 'Small Cap'),
		array('C', 'Small Cap')
	);

	protected $reits2 =
	array(
		array('A', 'REITs'),
		array('B', 'REITs'),
		array('C', 'Corps')
	);

	protected $thirdparty =
	array(
		array('*', 'Third Party Property Management Fee<br />% of Effective Gross Income')
	);

	protected $salescostaspercentsaleprice =
	array(
		array('*', 'Sales Cost as a % Sale Price')
	);

	protected $retention =
	array(
		array('*', ' % Retention')
	);

	protected $NoOfMonths =
	array(
		array('*', 'No. Of Months')
	);


	/* --- ROW ARRAYS --- */

	protected $property =
	array(
		array('',  1,1, 'OFFICE', FALSE),
		array('A', 1,1, 'Downtown Office Land', '129'),
		array('B', 1,1, 'Suburban Office Land', '129'),
		array('C', 1,1, 'Downtown Class "AA" Office', '129'),
		array('D', 1,1, 'Downtown Class "B" Office', '129'),
		array('E', 1,1, 'Suburban Class "A" Office', '129'),
		array('F', 1,1, 'Suburban Class "B" Office', '129'),
		array('',  1,1, 'RETAIL', FALSE),
		array('G', 1,1, 'Tier I Regional Mall', '129'),
		array('H', 1,1, 'Tier II Regional Mall', '129'),
		array('I', 1,1, 'Food Anchored Retail Strip', '129'),
		array('J', 1,1, 'Power Centre', '129'),
		array('K', 1,1, 'Community Mall', '129'),
		array('',  1,1, 'INDUSTRIAL', FALSE),
		array('L', 1,1, 'Multi-Tenant Industrial', '129'),
		array('M', 1,1, 'Single Tenant Industrial', '129'),
		array('N', 1,1, 'Industrial Land', '129'),
		array('O', 1,1, 'Multi-Unit Residential', '129'),
		array('P', 1,1, 'Hotel', '129')
		//array('Q', 1,1, 'REIT Units', FALSE)
	);

	protected $peracre =
	array(
		array('A', 10, 10, '$ Per Acre in Thousands', FALSE)
	);

	protected $psf =
	array(
		array('A', 10, 10, 'Net Face Rental Rate ($ psf)', FALSE)
	);

	protected $arr2 =
	array(
		array('A', 10,10, 'Net Effective Rate (($ psf) Base Building)', FALSE),
		array('B', 10,10, 'Face Rate (($ psf) Built-out from Base Building)', FALSE)
	);


	protected $arrTerm =
	array(
		array('A', 10,10, 'Next 12 Months', FALSE),
		array('B', 10,10, '13 - 24 Months', FALSE),
		array('C', 10,10, '25 - 36 Months', FALSE),
		array('D', 10,10, 'Stabilized Long Term', FALSE)
	);

	protected $arr4 =
	array(
		array('A', 10,10, 'Economic Face Rate ($ psf)', FALSE),
		array('B', 10,10, 'Development Yield', FALSE),
	);

	protected $arr5 =
	array(
		array('A', 10,10, '$ Per s.f. Buildable', FALSE),
		array('B', 1,1, 'Benchmark Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123'),
	);

	protected $arr6 =
	array(
		array('A', 10,10, 'Development Yield %', FALSE),
		array('B', 10,10, 'Stabilized Cap Rate<small><sup>1</sup></small>', FALSE),
	);

	protected $arrRate =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI before leasing costs)', FALSE),
		array('B', 10,10, 'Internal Rate of Return', FALSE),
		array('C', 10,10, 'Terminal Cap Rate', FALSE),
		array('D', 1,1, 'Benchmark Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123')
	);

	protected $arrRate_1 =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI)', FALSE),
		array('B', 10,10, 'Internal Rate of Return', FALSE),
		array('C', 10,10, 'Terminal Cap Rate', FALSE),
		array('D', 1,1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123')
	);

	protected $arrResidentialRate =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI)', FALSE),
		array('B', 10,10, '$ Price Per Suite', FALSE),
		array('C', 10,10, 'Internal Rate of Return (IRR)', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1,1, 'Benchmark Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123')
	);

	protected $arrMultiUnitResidentialRate =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI) Unlevered', FALSE),
		array('B', 10,10, '$ Price Per Suite', FALSE),
		array('C', 10,10, 'Internal Rate of Return (IRR) Unlevered', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1,1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123')
	);

	protected $arrRate_2 =
	array(
		array('7', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI)', FALSE),
	);

	protected $arrRate_3 =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1)', FALSE),
		array('B', 1,1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123')
	);

	protected $arrNewConstructionRate =
	array(
		array('A', 10,10, 'Overall Cap Rate', FALSE),
		array('B', 10,10, '$ Per s.f.', FALSE),
		array('C', 10,10, 'Internal Rate of Return (Unlevered)', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 10,10, 'Maximum Loan to Value Expectation', FALSE),
		array('F', 10,10, 'Minimum Spread over 10 Year Canada Bonds (bps) at Loan to Value indicated above', FALSE),
		array('G', 10,10, 'Internal Rate of Return (levered)', FALSE),
		array('H', 1,1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />Same / No&nbsp;Change = 3</small>', '123')
	);

	protected $arrTurnover =
	array(
		array('A', 10,10, '% Turnover', FALSE)
	);

	protected $arrCostperSuite =
	array(
		array('A', 10,10, 'Cost $ per Suite', FALSE)
	);

	protected $arrReturn =
	array(
		array('A', 10,10, 'Return (%)', FALSE)
	);

	protected $arrIncDecrease =
	array(
		array('A', 10,10, '% increase or decrease', FALSE)
	);

	protected $arrIncDecrease_1 =
	array(
		array('A', 1,1, 'Increase = 1<br />Decrease = 2<br />Same / No&nbsp;Change = 3', '123')
	);

	protected $arrRateSF =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI before leasing costs)', FALSE),
		array('B', 10,10, '$ Per s.f.', FALSE),
		array('C', 10,10, 'Internal Rate of Return', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1,1, 'Benchmark Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123'),
	);

	protected $arrRateSF_1 =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI before leasing costs)', FALSE),
		array('B', 10,10, '$ Per s.f.', FALSE),
		array('C', 10,10, 'Internal Rate of Return', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1,1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />Same / No&nbsp;Change = 3</small>', '123'),
	);

	protected $arrRateSF_2 =
	array(
		array('A', 10,10, 'Market Rent', FALSE),
		array('B', 10,10, 'Gross Sales Yr 1', FALSE),
		array('C', 10,10, 'Gross Sales Yr 2-5', FALSE),
		array('D', 10,10, 'Gross Sales Yr 6-10', FALSE),
		array('E', 10,10, 'Expenses ', FALSE),
	);

	protected $arrRateSF_3 =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI)', FALSE),
		array('B', 10,10, '$ Per Suite', FALSE),
		array('C', 10,10, 'Internal Rate of Return (IRR)', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1, 1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />Same / No&nbsp;Change = 3</small>', '123'),
	);

	protected $arrRateSF_4 =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI)', FALSE),
		array('B', 10,10, '$ Per s.f.', FALSE),
		array('C', 10,10, 'Internal Rate of Return', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1, 1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123'),
	);

	protected $arrRateSF_5 =
	array(
		array('A', 10,10, 'Overall Cap Rate (Yr 1 Stabilized NOI before leasing costs)', FALSE),
		array('B', 10,10, '$ Per s.f.', FALSE),
		array('C', 10,10, 'Internal Rate of Return', FALSE),
		array('D', 10,10, 'Terminal Cap Rate', FALSE),
		array('E', 1, 1, 'Value Trend next 12 months<br /><small>Increase = 1<br />Decrease = 2<br />No&nbsp;Change = 3</small>', '123'),
	);

	protected $arr7 =
	array(
		array('A', 10,10, '< 5', FALSE),
		array('B', 10,10, '5 - 10', FALSE),
		array('C', 10,10, '10 - 20', FALSE),
		array('D', 10,10, '20 - 50', FALSE),
		array('E', 10,10, '50 - 100', FALSE),
		array('F', 10,10, '> 100', FALSE)
	);

	protected $arr8 =
	array(
		array('A', 10,10, '$5 - $10 M', FALSE),
		array('B', 10,10, '$10 - $20 M', FALSE),
		array('C', 10,10, '$20 - $50 M', FALSE),
		array('D', 10,10, '$50 - $100 M', FALSE),
		array('E', 10,10, '> $100 M', FALSE)
	);

	protected $arrOfficeClasses =
		array(
			array('A', 1,1, 'Downtown Class "AA" Office', '123'),
			array('B', 1,1, 'Downtown Class "B" Office', '123'),
			array('C', 1,1, 'Suburban Class "A" Office', '123'),
			array('D', 1,1, 'Suburban Class "B" Office', '123'),
		);

	protected $arrClasses =
		array(
			array('A', 10,10, 'Downtown Class "AA"', FALSE),
			array('B', 10,10, 'Downtown Class "B"', FALSE),
			array('C', 10,10, 'Suburban Class "A"', FALSE),
			array('D', 10,10, 'Suburban Class "B"', FALSE),
		);

	protected $arrClasses_1 =
		array(
			array('A', 10,10, 'Downtown Class "AA" Office', FALSE),
			array('B', 10,10, 'Downtown Class "B" Office', FALSE),
			array('C', 10,10, 'Suburban Class "A" Office', FALSE),
			array('D', 10,10, 'Tier I Regional Mall', FALSE),
			array('E', 10,10, 'Multi-Tenant Industrial', FALSE),
			array('F', 10,10, 'Suburban Multiple Unit Residential', FALSE),
		);

	protected $arrInflation =
	array(
		array('A', 10,10, 'Market Rate Inflation', FALSE),
		array('B', 10,10, 'Expenses Inflation', FALSE),
	);

	protected $arrInflation_1 =
	array(
		array('A', 10,10, 'Market Rent', FALSE),
		array('B', 10,10, 'Expenses', FALSE),
	);

	protected $arr_rentalrate =
	array(
		array('A', 10,10, 'Yr 1-5 $ Psf', FALSE),
		array('B', 10,10, 'Yr 6-10 $ Psf', FALSE),
	);

	protected $arrBenchmark =
	array(
		array('A', 10,10, 'Tier I Regional Mall', FALSE),
		array('B', 10,10, 'Tier II Regional Mall', FALSE),
		array('C', 10,10, 'Community Mall', FALSE),
		array('D', 10,10, 'Power Centre', FALSE),
		array('E', 10,10, 'Food Anchored Retail Strip', FALSE)
	);

	protected $arrBudget =
	array(
		array('A', 10,10, 'Management (% of EGI)', FALSE),
		array('B', 10,10, 'Reserve (allowance) for replacement of capital items (% of EGI)', FALSE),
		array('C', 10,10, 'Repairs &amp; Maintenance (no major repairs) $ per suite', FALSE),
		array('D', 10,10, 'Staffing or wages / benefits, includes rental value of apartment ($ per suite)', FALSE),
		array('E', 10,10, 'Security $ per suite', FALSE),
		array('F', 10,10, 'Advertising $ per suite', FALSE)
	);

	protected $arrPropertyType =
	array(
		array('A', 10,10, 'Office', FALSE),
		array('B', 10,10, 'Retail', FALSE),
		array('C', 10,10, 'Industrial', FALSE),
		array('D', 10,10, 'Multi Residential', FALSE),
		array('E', 10,10, 'Hotel', FALSE),
		array('F', 10,10, 'Other - explain other', FALSE)
	);

	protected $arrCapital =
	array(
		array('A', 10,10, 'Maximum Loan to value ratio (1st Mortgage) (%)', FALSE),
		//array('B', 10,10, 'Minimum Dept Service Ratio (%)', FALSE),
		array('C', 10,10, 'Longest Amortization (Yrs)', FALSE),
		array('D', 10,10, 'Based on the answers provided above, what is the minimum spread over 5-year Canada Bonds, before fees (bps)', FALSE),
		array('E', 10,10, 'Based on the answers provided above, what is minimum  spread over 10-year Canada Bonds, before fees (bps)', FALSE),
		array('F', 10,10, 'What do consider to be the preferred term today (Yrs)', FALSE),
		array('G', 1, 1, 'Describe Current Availability of Non Recourse ?<br /><small>Poor = 1<br />Good = 2<br />Same / Excellent = 3</small>', '123'),
		array('H', 10,10, 'Additional BPs for Non Recourse debt (assuming Good to Excellent availability), based on preferred term.', FALSE)
		//array('H', 10,10, 'If Availability of Non Recourse is Good to Excellent, what is the additional bps cost to the preferred term indicated above?', FALSE)
	);

	protected $arrDiscountPremiumPercent =
	array(
		array('A', 1, 1, 'Discount or Premium?<br /><small>Discount = 1<br />Premium = 2</small>', '12'),
		array('C', 10,10, 'Insert percentage<br /><small>(i.e. 3% or 12% etc.)</small>', FALSE),
	);

	public function __construct(/* cString */ $name, /* cString */ $title)
	{
		parent::__construct(/* cString */ $name, /* cString */ $title);

		$this->market = $this->FilterMarket($this->market);
		$this->market_alt = $this->FilterMarket($this->market_alt);
		$this->market_example_1 = $this->FilterMarket($this->market_example_1);
		$this->market_example_2 = $this->FilterMarket($this->market_example_2);
		$this->market_example_3 = $this->FilterMarket($this->market_example_3);
		$this->market_example_4 = $this->FilterMarket($this->market_example_4);
		$this->market_example_5 = $this->FilterMarket($this->market_example_5);
		$this->market_example_6 = $this->FilterMarket($this->market_example_6);
		$this->market_example_7 = $this->FilterMarket($this->market_example_7);
		$this->market_example_8 = $this->FilterMarket($this->market_example_8);
		$this->market_example_9 = $this->FilterMarket($this->market_example_9);
		$this->market_example_10 = $this->FilterMarket($this->market_example_10);
		$this->market_example_11 = $this->FilterMarket($this->market_example_11);
		$this->market_example_12 = $this->FilterMarket($this->market_example_12);
		$this->market_example_13 = $this->FilterMarket($this->market_example_13);
		$this->market_example_14 = $this->FilterMarket($this->market_example_14);
		$this->market_example_15 = $this->FilterMarket($this->market_example_15);
		$this->market_example_16 = $this->FilterMarket($this->market_example_16);

		$this->property = $this->FilterProperty($this->property);
	}

	private function FilterMarket($market_in)
	{
		$user = $GLOBALS['user'];
		$split_array = split(',', $user->GetMARKETS());

		//echo 'user->GetMARKETS = ' . $user->GetMARKETS() . '<br />';

		$market_out = '';

		foreach ($split_array as $key => $value)
		{
			$v = (int)$value;
			$v--;

			//echo 'k = ' . $key . '<br />';
			//echo 'v = ' . $v . '<br />';

			$market_out[$key] = $market_in[$v];
		}

		return $market_out;
	}

	private function FilterProperty($property_in)
	{
		$user = $GLOBALS['user'];
		$split_array = split(',', $user->GetEXPERTISE());
		$property_out = '';
		$i=0;

		foreach ($split_array as $key => $value)
		{
			if ($value==1) /* OFFICE */
			{
				$property_out[$i] = $property_in[0]; $i++;
				$property_out[$i] = $property_in[1]; $i++;
				$property_out[$i] = $property_in[2]; $i++;
				$property_out[$i] = $property_in[3]; $i++;
				$property_out[$i] = $property_in[4]; $i++;
				$property_out[$i] = $property_in[5]; $i++;
				$property_out[$i] = $property_in[6]; $i++;
			}
			else if ($value==2) /* RETAIL */
			{
				$property_out[$i] = $property_in[7]; $i++;
				$property_out[$i] = $property_in[8]; $i++;
				$property_out[$i] = $property_in[9]; $i++;
				$property_out[$i] = $property_in[10]; $i++;
				$property_out[$i] = $property_in[11]; $i++;
				$property_out[$i] = $property_in[12]; $i++;
			}
			else if ($value==3) /* INDUSTRIAL */
			{
				$property_out[$i] = $property_in[13]; $i++;
				$property_out[$i] = $property_in[14]; $i++;
				$property_out[$i] = $property_in[15]; $i++;
				$property_out[$i] = $property_in[16]; $i++;
			}
			else if ($value==4) /* MULTI-RESIDENTIAL */
			{
				$property_out[$i] = $property_in[17]; $i++;
			}
		}
		$property_out[$i] = $property_in[18]; $i++; /* Hotel */
		//$property_out[$i] = $property_in[19]; $i++; /* REIT Units */
		return $property_out;
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Property(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->property, $this->market_alt);
	}

	protected function /* void */ DrawInputMatrix_Property(/* cString */ $name)
	{
		$this->head_title='Buy / Sell Preferences<br />Please Indicate:<br />Buy = 1, Sell = 2, No Comment / Not Familiar = leave blank';
		$this->head_type='Property Type';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->property, $this->market_alt, $this->market_example_1);
	}

/* ----- */

	protected function /* void */ DrawInputMatrix_Arr1(/* cString */ $name)
	{
		$this->head_title='Downtown Class "AA" Office Building';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRateSF, $this->market, $this->market_example_1);
	}

/* ----- */

	protected function /* void */ DrawInputMatrix_Arr1_B(/* cString */ $name, $title, $c)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		if ($c == 2)
			$this->DrawInputMatrix($name, $this->arrRateSF_1, $this->market, $this->market_example_10);
		else
			$this->DrawInputMatrix($name, $this->arrRateSF, $this->market, $this->market_example_1);
	}

	protected function /* void */ DrawInputMatrix_ArrNewConst(/* cString */ $name)
	{
		$this->head_title='Toronto New Construction<br />Downtown Class "AA" &amp; Suburban Class "A" Office Building';
		$this->head_type='PHYSICAL EXAMPLE';
		//$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrNewConstructionRate, $this->class_physical, '');
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr2(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arr2, $this->market);
	}

	protected function /* void */ DrawInputMatrix_Arr2(/* cString */ $name)
	{
		$this->head_title='Downtown Class "AA" Net Effective / Face Rates<br />New Tenants - 5 Year Terms';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arr2, $this->market, $this->market_example_1);
	}

	protected function /* void */ DrawInputMatrix_Arr2_B(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arr2, $this->market, $this->market_example_10);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_TermMarket(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrTerm, $this->market);
	}

	protected function /* void */ DrawInputMatrix_Arr3(/* cString */ $name)
	{
		$this->head_title='Downtown Class "AA" Face Rate Change<br />% Change (Positive or Negative)';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->market, $this->market_example_1);
	}

	protected function /* void */ DrawInputMatrix_Arr3_B(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->market, $this->market_example_10);
	}

	protected function /* void */ DrawInputMatrix_Arr3_C(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->market, $this->market_example_5);
	}

	protected function /* void */ DrawInputMatrix_Arr3_D(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->market, $this->market_example_14);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_CapitalBenchmark(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrCapital, $this->benchmark);
	}

	protected function /* void */ DrawInputMatrix_CapitalBenchmark(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type= 'Parameter';
		$this->head_example= '';
		$this->DrawInputMatrix($name, $this->arrCapital, $this->benchmark, '');
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr4(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arr4, $this->market);
	}

	protected function /* void */ DrawInputMatrix_Arr4(/* cString */ $name)
	{
		$this->head_title='Downtown Class "AA" Economic Face Rate &amp; Development Yield';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arr4, $this->market, $this->market_example_1);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_PropertyType(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrPropertyType, $this->arrPropertyTypeOD);
	}

	protected function /* void */ DrawInputMatrix_PropertyType(/* cString */ $name)
	{
		$this->head_title='Planned Investment In Canada - Ownership or Debt<br />By Property Type';
		$this->head_type='';

		$row_arr = $this->arrPropertyType;
		$example_arr = $this->arrPropertyTypeOD;
		$col_arr = array( 'Ownership', 'Debt');
		$col_digits = "";
		$row_digits = "";

		$col_prefix = "_";
		$row_suffix = "_";

		for ($p=0; $p<count($row_arr); $p++) { if ($row_arr[$p][4]) $row_digits=TRUE; }

		$this->GetContext()->Output('matrix_open.php');

			?><p align="center"><strong><?php echo $this->head_title; ?></strong></p>
			<table class="matrix"><tr><th><?php echo $this->head_type; ?></th><?php
		for ($m=0; $m<count($col_arr); $m++)
		{
			?>
				<th align="center" colspan="2"><?php echo $col_arr[$m]; ?></th>
			<?php
		}
			?></tr>
			<tr class="example"><th></th>
			<?php
		for ($m=0; $m<count($example_arr); $m++)
		{
			?>
			<th><small><em><?php echo $example_arr[$m][1]; ?></em></small></th>
			<?php
		}
		?></tr>
		<?php
		for ($p=0; $p<count($row_arr); $p++)
		{
			?>
			<tr align="center">
				<th><?php echo $row_arr[$p][3]; ?></th>
			<?php
			for ($m=0; $m<count($example_arr); $m++)
			{
				?>
				<td>
				<?php
				echo $this->InputTextBox($this->MakeCellName($name, $row_arr[$p][0], $example_arr[$m][0]));
				?></td>
				<?php
			}
			?>
			</tr>
			<?php
		}
		?>
		<tr align="center">
		<th>Total</th>
		<?php
		for ($m=0; $m<count($example_arr); $m++)
		{
			?>
			<td><label id="<?php echo 'Total'.$example_arr[$m][0]; ?>" ></label></td>
			<?php
		}
		?>
		</tr>
		</table>
		<script language="javaScript" type="text/javascript">
			InputCallback('MI108A1O');
		</script>
		<?php
		$this->GetContext()->Output('matrix_close.php');
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr5(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arr5, $this->market);
	}

	protected function /* void */ DrawInputMatrix_Arr5(/* cString */ $name)
	{
		$this->head_title='Downtown Land Parcel - Office Use';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arr5, $this->market, $this->market_example_2);
	}


/* ----- */


	protected function /* bool */ AddInputMatrix_Arr6(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRateSF, $this->pv_example_1);
	}

	protected function /* void */ DrawInputMatrix_Arr6(/* cString */ $name)
	{
		$this->head_title='Montreal Midtown Class "A" Office Building';
		$this->head_type='Physical Example';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrRateSF, $this->pv_example_1, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr6_3(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRateSF_3, $this->pv_example_2);
	}

	protected function /* void */ DrawInputMatrix_Arr6_3(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Physical Example';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrRateSF_3, $this->pv_example_2, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr7(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arr2, $this->pv_example_1);
	}

	protected function /* void */ DrawInputMatrix_Arr7(/* cString */ $name)
	{
		$this->head_title='Montreal Midtown Class "A" Office Building<br />Net Effective / Face Rates New Tenants - 5 Year Terms';
		$this->head_type='Physical Example';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arr2, $this->pv_example_1, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr8(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrTerm, $this->pv_example_1);
	}

	protected function /* void */ DrawInputMatrix_Arr8(/* cString */ $name)
	{
		$this->head_title='Montreal Midtown Class "A" Office Rate Change<br />% Change (Positive or Negative)';
		$this->head_type='Physical Example';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->pv_example_1, null);
	}

	protected function /* void */ DrawInputMatrix_Arr8_B(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type= 'Physical Example';
		$this->head_example= '';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->pv_example_1, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Arr9(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arr7, $this->thirdparty);
	}

	protected function /* void */ DrawInputMatrix_Arr9(/* cString */ $name)
	{
		$this->head_title='Third Party Property Management Fees<br />(Office Property)';
		$this->head_type='Asset Value ($ Million)';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arr7, $this->thirdparty, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_VacancyBarometer(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrOfficeClasses, $this->market);
	}

	protected function /* bool */ AddInputMatrix_VacancyBarometer_B(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrClasses, $this->market);
	}

	protected function /* void */ DrawInputMatrix_VacancyBarometer(/* cString */ $name)
	{
		$this->head_title='Altus - InSite Vacancy Barometer<br />Next 3 Months<br />code 1 = increase, code 2 = decrease, code 3 = no change';
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrOfficeClasses, $this->market, $this->market_example_1);
	}

	protected function /* void */ DrawInputMatrix_VacancyBarometer_B(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrClasses, $this->market, $this->market_example_1);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_arrRatemarket(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRate, $this->market);
	}

	protected function /* bool */ AddInputMatrix_arrRatemarket_1(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRate_1, $this->market);
	}

	protected function /* bool */ AddInputMatrix_arrRatemarket_2(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRate_2, $this->market);
	}

	protected function /* void */ DrawInputMatrix_arrRatemarket_2(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRate_2, $this->market, $this->market_example_9);
	}

	protected function /* void */ DrawInputMatrix_arrRatemarket_1(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRate_1, $this->market, $this->market_example_11);
	}

	protected function /* void */ DrawInputMatrix_arrRatemarket_12(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRate_1, $this->market, $this->market_example_12);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Tier2RegionalMall(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRate, $this->province);
	}

	protected function /* void */ DrawInputMatrix_Tier2RegionalMall(/* cString */ $name)
	{
		$this->head_title='Tier II Regional Mall';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRate, $this->province, $this->province_example_1);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_Tier2RegionalMallInflation(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrInflation, $this->province);
	}

	protected function /* void */ DrawInputMatrix_Tier2RegionalMallInflation(/* cString */ $name)
	{
		$this->head_title='Tier II Regional Mall Rent &amp; Expense Annual Percent Change';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrInflation, $this->province, $this->province_example_1);
	}

/* ----- */

	protected function /* bool */ AddInputPowerCentreRentExpense (/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrInflation_1, $this->market);
	}

	protected function /* void */ DrawInputPowerCentreRentExpense(/* cString */ $name, $title)
	{
		$this->head_title=$title;
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrInflation_1, $this->market, null);
	}

/* ----- */
	protected function /* void */ DrawInputMatrix_CommunityMall(/* cString */ $name)
	{
		$this->head_title='Community Mall';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRate, $this->market, $this->market_example_4);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_tenantRetention(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrBenchmark, $this->retention);
	}

	protected function /* void */ DrawInputMatrix_tenantRetention(/* cString */ $name)
	{
		$this->head_title='Retail Tenant Retention Ratio<br />CRU % Renew Only';
		$this->head_type='Benchmark Property';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrBenchmark, $this->retention, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_LagVacancy(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrBenchmark, $this->NoOfMonths);
	}

	protected function /* void */ DrawInputMatrix_LagVacancy(/* cString */ $name)
	{
		$this->head_title='Retail Lag Vacancy<br />(Including Fixturing Period)';
		$this->head_type='Benchmark Property';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrBenchmark, $this->NoOfMonths, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_VacancyAndCredit(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrBenchmark, $this->market);
	}

	protected function /* void */ DrawInputMatrix_VacancyAndCredit(/* cString */ $name)
	{
		$this->head_title='Retail Vacancy &amp; Credit Allowance<br />% of CRU Income';
		$this->head_type='Benchmark Property';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrBenchmark, $this->market, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_arrRateSF(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRateSF, $this->market);
	}

	protected function /* bool */ AddInputMatrix_arrRateSF_1(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRateSF_1, $this->market);
	}

	protected function /* bool */ AddInputMatrix_arrRateSF_2(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrRateSF_2, $this->market);
	}

	protected function /* void */ DrawInputMatrix_Arr6_2(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type= 'Physical Example';
		$this->head_example= '';
		$this->DrawInputMatrix($name, $this->arrRateSF_2, $this->market, null);
	}

	protected function /* bool */ AddInputMatrix_arrNewConClass(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrNewConstructionRate, $this->class_physical);
	}

	protected function /* void */ DrawInputMatrix_SingleIndustrial(/* cString */ $name)
	{
		$this->head_title='Single-Tenant Industrial Building';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRateSF, $this->market, $this->market_example_5);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_arr_rentalrate(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arr_rentalrate, $this->market);
	}

	protected function /* void */ DrawInputMatrix_arr_rentalrate(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type= 'Physical Example';
		$this->head_example= 'Physical Example';
		$this->DrawInputMatrix($name, $this->arr_rentalrate, $this->market, $this->market_example_5);
	}

/* ----- */

	protected function /* void */ DrawInputMatrix_MultiIndustrial(/* cString */ $name)
	{
		$this->head_title='Multi-Tenant Industrial Building';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrRateSF, $this->market, $this->market_example_6);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_InitialNetFace(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->psf, $this->market);
	}

	protected function /* void */ DrawInputMatrix_InitialNetFace(/* cString */ $name)
	{
		$this->head_title='Initial Net Face Rental Rate - Multi-Tenant Industrial';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->psf, $this->market, $this->market_example_7);
	}

/* ----- */

	protected function /* void */ DrawInputMatrix_NetFaceGrowth(/* cString */ $name)
	{
		$this->head_title='Net Face Rental Rate Change - Multi-Tenant Industrial<br />% Change (Positive or Negative)';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrTerm, $this->market, $this->market_example_8);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_SuburbanMulti(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrResidentialRate, $this->market);
	}

	protected function /* void */ DrawInputMatrix_SuburbanMulti(/* cString */ $name)
	{
		$this->head_title='Suburban Multi-Unit Residential';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrResidentialRate, $this->market, $this->market_example_9);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_arrBudget(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrBudget, $this->market);
	}

	protected function /* void */ DrawInputMatrix_arrBudget(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrBudget, $this->market, $this->market_example_9);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_arrDebtCost(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrIncDecrease_1, $this->arrType);
	}

	protected function /* void */ DrawInputMatrix_arrDebtCost(/* cString */ $name, $title)
	{
		$this->head_title= $title;
		$this->head_type='Type';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrIncDecrease_1, $this->arrType, '');
	}

/* ----- */
	protected function /* bool */ AddInputMatrix_SuburbanMultiAnnual(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrTurnover, $this->market);
	}

	protected function /* void */ DrawInputMatrix_SuburbanMultiAnnual(/* cString */ $name)
	{
		$this->head_title='Suburban Multi-Unit Residential<br />(Annual Tenant Turnover)';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrTurnover, $this->market, $this->market_example_9);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_SuburbanMultiSuiteCost(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrCostperSuite, $this->market);
	}

	protected function /* void */ DrawInputMatrix_SuburbanMultiSuiteCost(/* cString */ $name)
	{
		$this->head_title='Suburban Multi-Unit Residential<br />(Suite Turnover Cost)';
		$this->head_type='Type';
		$this->head_example='Physical Example';
		$this->DrawInputMatrix($name, $this->arrCostperSuite, $this->market, $this->market_example_9);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_RequiredReturn(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrReturn, $this->reits);
	}

	protected function /* void */ DrawInputMatrix_RequiredReturn(/* cString */ $name)
	{
		$this->head_title='Total Required Return';
		$this->head_type='';
		$this->head_example='';
		$this->DrawInputMatrix($name, $this->arrReturn, $this->reits, null);
	}

/* ----- */

	protected function /* bool */ AddInputMatrix_PricePerformance(/* cString */ $name, /* cString */ $label)
	{
		return $this->AddInputMatrix($name, $label, $this->arrIncDecrease, $this->caps);
	}

	protected function /* void */ DrawInputMatrix_PricePerformance(/* cString */ $name)
	{
		$this->head_title='&nbsp;';
		$this->head_type='&nbsp;';
		$this->head_example='Performance';
		$this->DrawInputMatrix($name, $this->arrIncDecrease, $this->caps, $this->reits2);
	}

/* ===== */

	private function MakeCellName($name, $row, $col)
	{
		/* First deal with exceptions */

		if ($name=='CC48' && $row=='A')
		{
			     if ($col == 'A') return $name . 'A_R';
			else if ($col == 'B') return $name . 'B_R';
			else if ($col == 'C') return $name . 'B_C';
		}
		else if ($name=='CC43' && $row=='A')
		{
			     if ($col == 'A') return $name . 'A';
			else if ($col == 'B') return $name . 'B';
		}
		else if ($name=='VI133' && $row=='A')
		{
			return $name . '_' . $col;
		}
		else if ($name=='VR139' && $row=='A')
		{
			return $name . '_' . $col;
		}
		else if ($name=='VM118' && $row=='A')
		{
			return $name . '_' . $col;
		}
		else if ($name=='VM82' && $row=='A')
		{
			return $name . '_' . $col;
		}
		// ------------- Added by B.D. Feb 25, 2009
		else if ($name=='CD85A')
		{
			return $name . '' . $col;
		}
		else if ($name=='MI108')
		{
			return $name . '' . $row . $col;
		}
		// ----------------------------------------

		$col_suffix = '';
		if ($col<>'*') $col_suffix = '_' . $col;

		$cell_name = $name . '' . $row . $col_suffix;

//		echo $cell_name . ' --<br/>';

		return $cell_name;
	}

	protected function /* bool */ AddInputMatrix(/* cString */ $name, /* cString */ $label, $row_arr, $col_arr)
	{
		$ret = true;

		for ($p=0; $p<count($row_arr); $p++)
		{
			if ($row_arr[$p][0]=='') continue;

			for ($m=0; $m<count($col_arr); $m++)
			{
				$cell_name=$this->MakeCellName($name, $row_arr[$p][0], $col_arr[$m][0]);
				$cell_label = $label . ' (' . $row_arr[$p][3] . ', ' . $col_arr[$m][1] . ')';

				$ret &= $this->AddInput(new cInputTextBox($cell_name, $cell_label, ''));

				$this->FindInput($cell_name)->SetSize($row_arr[$p][1]);

				$this->FindInput($cell_name)->SetMaxlength($row_arr[$p][2]);
			}
		}

		return $ret;
	}

	protected function /* void */ DrawInputMatrix(/* cString */ $name, $row_arr, $col_arr, $example_arr)
	{
		$col_digits = "";
		$row_digits = "";

		$col_prefix = "_";
		$row_suffix = "_";

		if ($name=="Q1") $col_digits = "129";
		//else if ($name=="MI103") $col_digits = "123";
		else if ($name=="VO103") $col_digits = "123";
		else if ($name=="VO119") $row_suffix = "";
		//--- added by BD Feb 25, 2009
		else if ($name=="VM127") $row_suffix = "";
		//----------
		else if ($name=="CD85A") $row_suffix = "";

		for ($p=0; $p<count($row_arr); $p++) { if ($row_arr[$p][4]) $row_digits=TRUE; }

		$this->GetContext()->Output('matrix_open.php');
	?>

	<p align="center"><strong><?php echo $this->head_title; ?></strong></p>

	<table class="matrix">

	<tr>
		<th><?php echo $this->head_type; ?></th>
	<?php
	for ($m=0; $m<count($col_arr); $m++)
	{
	?>
		<th align="center"><?php echo $col_arr[$m][1]; if ($col_digits) $this->Digits('<br />', $col_prefix . $col_arr[$m][0], $col_digits); ?></th>
	<?php
	}

	if ($row_digits) echo '<th><small class="quickselect_head">SELECT ALL<br />FOR ROW</small></th>';

	?>
	</tr>

	<?php
	if ($this->head_example<>'')
	{
	?>
	<tr class="example">
		<th><small><em><?php echo $this->head_example; ?></em></small></th>
	<?php
	for ($m=0; $m<count($col_arr); $m++)
	{
	?>
		<th><small><em><?php echo $example_arr[$m][1]; ?></em></small></th>
	<?php
	}

	if ($row_digits) echo '<th>&nbsp;</th>';

	?>
	</tr>
	<?php
	}
	?>

	<?php
	for ($p=0; $p<count($row_arr); $p++)
	{
	?>

	<tr align="center">
		<th><?php echo $row_arr[$p][3]; ?></th>

		<?php

		$row_digits_tmp=$row_arr[$p][4];

		if ($row_arr[$p][0]=='')
		{
			$n = count($col_arr);
			echo "<td colspan='" . $n . "'>&nbsp;</td>";
			if ($row_digits) $row_digits_tmp=FALSE;
		}
		else for ($m=0; $m<count($col_arr); $m++)
		{
		?>
			<td><?php echo $this->InputTextBox($this->MakeCellName($name, $row_arr[$p][0], $col_arr[$m][0])); ?></td>
		<?php
		}

		if ($row_digits_tmp) { echo '<th align="center">'; $this->Digits('', $row_arr[$p][0] . $row_suffix, $row_digits_tmp); echo '</th>'; }
		else if ($row_digits) echo '<th>&nbsp;</th>';
		?>
	</tr>

	<?php
	}

	if ($name=='Q1')
	{
		$n = count($col_arr);
		echo '<tr align="center">';
		echo '<th>REIT Units</th>';
		echo '<td colspan="' . $n . '">';
		echo $this->InputTextBox($name . 'Q');
		echo '</td>';
		if ($row_digits) echo '<th>&nbsp;</th>';
		echo '</tr>';
	}

	?>

	</table>

	<?php
		$this->GetContext()->Output('matrix_close.php');
	}

	public function /* void */ Digits($br, $code, $type)
	{
		//echo $br;

		echo '<div class="quickselect_digits">';
		echo '<img src="images/digits/1.gif" width="13" height="13" alt="1" hspace="1" vspace="1" onclick="setCol(1,\'' . $code . '\');" onkeypress="this.click();" />';
		echo '<img src="images/digits/2.gif" width="13" height="13" alt="2" hspace="1" vspace="1" onclick="setCol(2,\'' . $code . '\');" onkeypress="this.click();" />';

		if ($type=="123")
		{
			echo '<img src="images/digits/3.gif" width="13" height="13" alt="3" hspace="1" vspace="1" onclick="setCol(3,\'' . $code . '\');" onkeypress="this.click();" />';
		}
		else
		{
			//echo '<img src="images/digits/9.gif" width="13" height="13" alt="9" hspace="1" vspace="1" onclick="setCol(9,\'' . $code . '\');" onkeypress="this.click();" />';
		}

		echo '</div>';
	}

/* ===== */


	public function /* void */ ParentDisplay()
	{
		//parent::Display();

		/* cInt */ $i = 0;

		//$this->Context->DisplayFormBefore($this);
		while ($input = $this->GetInput($i))
		{
			$this->DisplayInput($input, true);
			$i++;
		}
		//$this->Context->DisplayFormAfter($this);

	}

	public function /* void */ Display()
	{
		$this->GetContext()->DisplaySheetBefore($this);
		$this->GetContext()->Output('displ.cSheet_survey.php');
		$this->GetContext()->DisplaySheetAfter($this);
	}

	protected function /* void */ InputTextBox($name)
	{
		//echo $name . '<br/>';
		$this->GetContext()->ExecTemplateByFilename('class.InputTextBox_Matrix.php', $this->FindInput($name));
	}


/* ===== */

	private function isInteger($input)
	{
		return(ctype_digit(strval($input)));
	}

	private function /* cBool */ Result(/* cBool */ $validation, /* cInput */ $input, /* cString */ $warning)
	{
		if ((!$validation) && (strlen($warning) > 0)) $input->SetWarning($warning);
		return $validation;
	}

	protected function /* cBool */ Validate_1_2_x_or_empty(/* cInput */ $input, /* cString */ $warning, $x, $filter)
	{
		$value = trim($input->GetValue());

		if (strlen($value)<1) return true;

		$ret = false;
		$name = trim($input->GetName());

		if (strlen($filter)>0)
		{
			if (strstr($name, $filter)==FALSE) return true;
		}

		if ($this->isInteger($value))
		{
			if (($value == 1) || ($value == 2) || ($value == $x))
			{
				$ret = true;
			}
		}

		return $this->Result($ret, $input, $warning);
	}


	protected function /* cBool */ Validate_Rate(/* cInput */ $input, /* cString */ $warning, $filter)
	{
		$value = trim($input->GetValue());

		if (strlen($value)<1) return true;

		$name = trim($input->GetName());

		if (strlen($filter)>0)
		{
			if (strstr($name, $filter)==FALSE) return true;
		}

		$validator = new cValidator(true);

		return $validator->ValidateNumericRange($input, 1, 100, $warning);
	}

	protected function /* cBool */ Validate_Change(/* cInput */ $input, /* cString */ $warning, $filter)
	{
		$value = trim($input->GetValue());

		if (strlen($value)<1) return true;

		$name = trim($input->GetName());

		if (strlen($filter)>0)
		{
			if (strstr($name, $filter)==FALSE) return true;
		}


		$validator = new cValidator(true);

		return $validator->ValidateNumericRange($input, -100, 100, $warning);

	}


	protected function /* cBool */ Validate_Large(/* cInput */ $input, /* cString */ $warning, $filter)
	{
		$value = trim($input->GetValue());

		if (strlen($value)<1) return true;

		$name = trim($input->GetName());

		if (strlen($filter)>0)
		{
			if (strstr($name, $filter)==FALSE) return true;
		}

		$validator = new cValidator(true);

		return $validator->ValidateNumericRange($input, 10000, 1000000, $warning);
	}

	protected function /* cBool */ Validate_Ignored($ret)
	{
		if (!$ret) return $ret;

		$db_array = array('dbtype' => 'mssql', 'server' => _DB_SERVER, 'database' => _DB_DATABASE, 'username' => _DB_USERNAME, 'password' => _DB_PASSWORD);

		$DB = new cDB;
		$DB->connect($db_array['server'], $db_array['username'], $db_array['password']);
		$DB->select_db($db_array['database']);

		$ret = true;
		$n = $this->GetInputCount();

		for ($i=0; $i<$n; $i++)
		{
			$input=$this->GetInput($i);
			$name = trim($input->GetName());
			$value = trim($input->GetValue());

			//echo $name . '<br />';

			$name_escaped = $DB->escape_string($name);
			$sql = "SELECT Ignored FROM dat_Question WHERE LabelShort = '" . $name_escaped . "'";

			//echo $sql;

			$result = $DB->query($sql);

			if ($line = $DB->fetch_array($result))
			{
				$ignored = $line['Ignored'];
				//echo 'Ignored value for field ' . $name . ' is ' . $ignored . '<br />';

				if ($value != $ignored) continue;

				$ret = false;
				$this->Result($ret, $input, 'THE VALUE ' . $ignored . ' IS NOT ALLOWED');
			}
			else
			{
				$ret = false;
				$this->Result($ret, $input, 'NOT FOUND IN THE DATABASE');
			}
		}


		$DB->close();

		return $ret;
	}
}
?>
