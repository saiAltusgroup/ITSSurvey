<%

' DK, Aug 01, 2008 - version 1.0

Class cTableFrequency

	Public TableObj
	Public CurrPeriodID
	Private LabelArr

	' Constructor
	Private Sub Class_Initialize()
		'Set TableObj=new cTableListbox
		Set TableObj=new cTableDatagrid

		CurrPeriodID=g_CURRENT_PERIOD_ID

		ReDim LabelArr(101)

		' Chart title | Q107 | Q108 | Q407 | Q307 | Q207

		LabelArr(0) = "Altus InSite Location Barometer|X|X|X|X|X"
		LabelArr(1) = "Altus InSite Property Type Barometer|X|X|X|X|X"
		LabelArr(2) = "Altus InSite Product/Market Barometer|X|X|X|X|X"
		LabelArr(3) = "Office Realty Commissions|||||X"
		LabelArr(4) = "Office Tenant Retention Ratio|X|X|||"
		LabelArr(5) = "Office Lag Vacancy||||X|"
		LabelArr(6) = "CBD Class ""AA"" Net Effective/Face Rates|||X||X"
		LabelArr(7) = "CBD Class ""AA"" Face Rate Inflation|||X||X"
		LabelArr(8) = "CBD Class ""B"" Office Valuation Parameters|||||X"
		LabelArr(9) = "CBD Land Parcel - Office Use|||X||X"
		LabelArr(10) = "Leasehold Downtown Class ""AA"" Valuation Parameters|X|X|||"
		LabelArr(11) = "Downtown ""AA"" Economic Face Rate & Development Yield|||X||"
		LabelArr(12) = "Suburban Class ""A"" Office Valuation Parameters||||X|"
		LabelArr(13) = "Suburban ""A"" Net Effective/Face Rates||||X|"
		LabelArr(14) = "Suburban ""A"" Face Rate Inflation||||X|"
		LabelArr(15) = "Suburban ""A"" Economic Face Rate & Development Yield||||X|"
		LabelArr(16) = "Suburban Office Land Parcel||X||X|"
		LabelArr(17) = "Downtown Class AA Office Valuation Parameters|X|X|X|X|X"
		LabelArr(18) = "Most Active Buyers and Sellers of Office Properties ||||X|"
		LabelArr(19) = "Montreal Midtown Class ""A"" Office Valuation Parameters (MONTREAL ONLY)|X|X|X|X|X"
		LabelArr(20) = "Montreal Midtown Class ""A"" Office Face Rate Inflation (MONTREAL ONLY)|X|X|X|X|X"
		LabelArr(21) = "Suburban Class ""B"" Net Effective / Face Rates|X|X|||"
		LabelArr(22) = "Suburban Class ""B"" Face Rate Inflation|X|X|||"
		LabelArr(23) = "Office Inducement & Realty Commission Inflation - Downtown Class ""AA"" Office|||||X"
		LabelArr(24) = "Suburban Class ""B"" Office Valuation Parameters|X|X|||"
		LabelArr(25) = "Third Party Property Management Fees (% of Effective gross income)|||X||"
		LabelArr(26) = "Third Party Asset Management Fees (% of Asset Value)|||X||"
		LabelArr(27) = "Toronto New Construction - Downtown Class ""AA"" & Suburban Class ""A"" Office Bldg|X|X|||"
		LabelArr(28) = "Motivation for Office Investment ||||X|"
		LabelArr(29) = "Discount Rate to amortize tenant concessions, free rent or other forms of inducements|||X||"
		LabelArr(30) = "Discount Rate to amortize tenant concessions, free rent or other forms of inducements|||X||"
		LabelArr(31) = "Space Adjustments at Renewal|X||||"
		LabelArr(32) = "CBD Class ""B"" Net Effective/Face Rates|||||X"
		LabelArr(33) = "CBD Class ""B"" Net Effective/Face Rates|||||X"
		LabelArr(34) = "CBD Class ""B"" Face Rate Inflation|||||X"
		LabelArr(35) = "Ground Lease Downtown Class ""AA"" Office Valuation Parameters|||||X"
		LabelArr(36) = "Gross Rent to Sales Ratio ||||X|"
		LabelArr(37) = "Enclosed Community Mall Valuation Parameters|||X||X"
		LabelArr(38) = "Retail Management Fees - % Eff Gross Inc|||||X"
		LabelArr(39) = "Single-Tenant Retail Valuation Parameters||||X|"
		LabelArr(40) = "Motivation for Retail Investment |X|X|||"
		LabelArr(41) = "Retail Lag Vacancy - Excluding Fixturing Period |||X||"
		LabelArr(42) = "Retail Lag Vacancy - Fixturing Period |||X||"
		LabelArr(43) = "Retail Credit Allowance|||X||"
		LabelArr(44) = "Regional Mall Inflation|X|X|||"
		LabelArr(45) = "Power Centre Valuation Parameters|X|X||X|"
		LabelArr(46) = "Power Centre Valuation Parameters|X|X||X|"
		LabelArr(47) = "Power Centre Valuation Parameters|X|X||X|"
		LabelArr(48) = "Power Centre Valuation Parameters|X|X||X|"
		LabelArr(49) = "Food Anchored Retail Strip Valuation Parameters|X|X||X|"
		LabelArr(50) = "Tier I Regional Mall Valuation Parameters|X|X|X|X|X"
		LabelArr(51) = "Tier II Regional Mall  Valuation Parameters|||X||X"
		LabelArr(52) = "Tier II Regional Mall Inflation |||X||"
		LabelArr(53) = "Most Active - Buyers & Sellers of Retail Properties||||X|"
		LabelArr(54) = "Retail Tenant Retention Ratio |||X||"
		LabelArr(55) = "Power Centre Inflation|X|X|||"
		LabelArr(56) = "Enclosed Community Mall Inflation|||||X"
		LabelArr(57) = "Food Anchored Retail Strip Inflation|X|X|||"
		LabelArr(58) = "Vacant Land Value for acquisition of similar Power Centre site||||X|"
		LabelArr(59) = "Multi-Tenant Industrial Valuation Parameters|X|X|X|X|X"
		LabelArr(60) = "Single-Tenant Industrial Valuation Parameters|X|X|X|X|X"
		LabelArr(61) = "Rental Rate Single Tenant Industrial|X|X||X|"
		LabelArr(62) = "Rental Rate Inflation - Single Tenant Industrial|X|X||X|"
		LabelArr(63) = "Initial Net Rental Rate - Multi-Tenant Industrial|||X||X"
		LabelArr(64) = "Rental Rate Inflation - Multi-Tenant Industrial|||X||X"
		LabelArr(65) = "Most Active - Buyers & Sellers of Industrial Properties|||X||"
		LabelArr(66) = "Multi-Unit Residential Valuation Parameters |X|X|X|X|X"
		LabelArr(67) = "Multi-Unit Residential Rental Building Expenses|X|X|||"
		LabelArr(68) = "Multi-Residential Construction|||||X"
		LabelArr(69) = "Multi-Residential Development Yield & Cap Rate|||||X"
		LabelArr(70) = "Multi-Residential Rental Inflation||||X|"
		LabelArr(71) = "Downtown Multi-Unit Residential Valuation Parameters ||||X|"
		LabelArr(72) = "Suburban Multi-Unti Residential (Suite Turnover Cost)|||X||"
		LabelArr(73) = "Buyer/Seller of Benchmark Multi-Unit Residential Leasehold|X|X|||"
		LabelArr(74) = "Multi-Unit Residential Leasehold Parameters |X|X|||"
		LabelArr(75) = "Annual Tenant Turnover |||X||"
		LabelArr(76) = "Multi-Unit Residential Rental - Renovation Cost|X|X|||"
		LabelArr(77) = "Midtown Multi-Unit Residential Valuation Parameters - Montreal |X|X||X|"
		LabelArr(78) = "Threats to Multi-Residential Investment ||||X|"
		LabelArr(79) = "Motivation for Multi Residential Investment |||X||"
		LabelArr(80) = "Estimated Market Value of Portfolio|X||||"
		LabelArr(81) = "Captial Sources - Total Required Return (REIT's & Publicly Traded Corp)|||X||X"
		LabelArr(82) = "Real Estate Equity Available|||X||X"
		LabelArr(83) = "Real Estate Equity Price Performance|||X||X"
		LabelArr(84) = "Three Most Active sources of First Mortgage Debt Capital|X|X||X|"
		LabelArr(85) = "Conventional Debt Market Parameters|X|X||X|"
		LabelArr(86) = "Debt Cost Outlook|X|X||X|"
		LabelArr(87) = "Structural Reserve |X|X|||"
		LabelArr(88) = "Reversion Terminal Cap Rate|X|X|||X"
		LabelArr(89) = "Recoverable Depreciation||||X|"
		LabelArr(90) = "Interest Rate Impact on Capitalization Rates|||||X"
		LabelArr(91) = "Deduct Selling Costs?|||||X"
		LabelArr(92) = "Total Sales Cost at Reversion|||||X"
		LabelArr(93) = "Partial Interests Property Mgmt Contract Not Included for a 20%, 33.3%, 50% Interest|X||||"
		LabelArr(94) = "Financial Indicators for CBD Class ""AA"" |||||X"
		LabelArr(95) = "Land / Valuation Parameters Split - 10 Year Old Valuation Parameters|||X||"
		LabelArr(96) = "Land / Valuation Parameters Split - 30 Year Old Valuation Parameters|||X||"
		LabelArr(97) = "Marketing Time - No. of Months||||X|"
		LabelArr(98) = "Insite-Altus Office Vacancy Barometer |X|X|X|X|X"
		LabelArr(99) = "Management Fee Profits|||X||"
		LabelArr(100) = "Planned Investment in Canada - Ownership & Debt|X|X|||"
		LabelArr(101) = "Homogeneous Portfolio Effect (Discount or Premium)|||||X"

	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub


	Public Function Read()

		Read = TableObj.TableOpen(0)

		Read = Read & TableObj.RowOpen & TableObj.HeadSpan(Lang("qtr_frequency"), 4) & TableObj.RowClose

		TableObj.RowID=Lang("topics")
		Read = Read & TableObj.RowOpen & TableObj.Head("Q1") & TableObj.Head("Q2") & TableObj.Head("Q3") & TableObj.Head("Q4") & TableObj.RowClose

		for i=0 to ubound(LabelArr)
			s=split(LabelArr(i), "|")
			TableObj.RowID=s(0)

			Read = Read & TableObj.RowOpen & TableObj.Cell(s(1)) & TableObj.Cell(s(5)) & TableObj.Cell(s(4)) & TableObj.Cell(s(3)) & TableObj.RowClose
		next

		Read = Read & TableObj.TableClose

	End Function

End Class

%>
