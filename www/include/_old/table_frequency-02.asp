<%

' DK, Aug 20, 2008 - version 1.0

Class cTableFrequency

	Private TableObj
	Private LabelArr

	' Constructor
	Private Sub Class_Initialize()

		Set TableObj=new cTableDatagrid
		TableObj.ReadOnly=True


		' RowType | ChartTitle | Q1 | Q2 | Q3 | Q4

		If gLang="fre" Then

		ReDim LabelArr(103)
		LabelArr(0) = "1|Préférences par produit||||"
		LabelArr(1) = "0|Baromètre Altus InSite par ville|X|X|X|X"
		LabelArr(2) = "0|Baromètre Altus InSite par genre de propriété|X|X|X|X"
		LabelArr(3) = "0|Baromètre Altus InSite par produit/marché|X|X|X|X"
		LabelArr(4) = "1|Bureaux||||"
		LabelArr(5) = "0|Commissions sur la location d'espace de bureaux||X||"
		LabelArr(6) = "0|Taux de rétention des locataires de bureaux|X|||"
		LabelArr(7) = "0|Délai d'inoccupation bureaux|||X|"
		LabelArr(8) = "0|Loyers nets effectifs/Loyers nets demandés - Édifices de bureaux classe AA au centre-ville||X||X"
		LabelArr(9) = "0|Inflation loyers nets demandés - Édifices de bureaux classe AA au centre-ville||X||X"
		LabelArr(10) = "0|Critères de rendement - Édifices de bureaux classe B au centre-ville||X||"
		LabelArr(11) = "0|Terrain pour bureaux au centre-ville||X||X"
		LabelArr(12) = "0|Critères de rendement - Édifice classe AA au centre-ville en tenure à bail|X|||"
		LabelArr(13) = "0|Loyer net économique et rendement sur le développement - Édifices de bureaux classe AA au centre-ville||||X"
		LabelArr(14) = "0|Critères de rendement - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(15) = "0|Loyers nets effectifs/loyers nets demandés - Édifice de bureaux classe A en banlieue|||X|"
		LabelArr(16) = "0|Inflation loyers nets demandés - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(17) = "0|Loyer net économique et rendement sur le développement - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(18) = "0|Terrain pour bureaux en banlieue|||X|"
		LabelArr(19) = "0|Critères de rendement - Édifices de bureaux classe AA au centre-ville|X|X|X|X"
		LabelArr(20) = "0|Acheteurs et vendeurs les plus actifs - Marché des édifices de bureaux|||X|"
		LabelArr(21) = "0|Critères de rendement - Édifices classe A en périphérie du centre-ville de Montréal (Montréal seulement)|X|X|X|X"
		LabelArr(22) = "0|Inflation loyers nets demandés - Édifices classe A en périphérie du centre-ville de Montréal (Montréal seulement)|X|X|X|X"
		LabelArr(23) = "0|Loyers nets effectifs/loyers nets demandés - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(24) = "0|Inflation loyers nets demandés - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(25) = "0|Inflation des incitations et commissions immobilières  - Édifices de bureaux classe AA au centre-ville||X||"
		LabelArr(26) = "0|Critères de rendement - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(27) = "0|Honoraires de gestion par un tiers (% du revenu brut effectif)||||X"
		LabelArr(28) = "0|Honoraires de gestion par un tiers (% de la valeur de l'actif sous gestion)||||X"
		LabelArr(29) = "0|Nouveaux projets à Toronto - Édifices classe AA au centre-ville et édifices classe B en banlieue|X|||"
		LabelArr(30) = "0|Motifs d'investissement sur le marché des bureaux|||X|"
		LabelArr(31) = "0|Escompte (%) sur amortissement des privilèges aux locataires, loyers gratuits et autres formes d'incitations locatives||||X"
		   LabelArr(32) = "0|Ajustement de l'aire locative au renouvellement|X|||"
		LabelArr(33) = "0|Loyers nets effectifs/Loyers nets demandés - Édifices classe B au centre-ville||X||"
		LabelArr(34) = "0|Inflation loyers nets demandés - Édifices de classe B au centre-ville||X||"
		LabelArr(35) = "0|Critères de rendement - Bail à construction pour un édifice de bureaux classe AA au centre-ville||X||"
		LabelArr(36) = "0|Ratio loyer brut vs ventes|||X|"
		LabelArr(37) = "0|Baromètre Altus-InSite sur l'inoccupation de bureaux|X|X|X|X"
		LabelArr(38) = "1|Centres commerciaux||||"
		LabelArr(39) = "0|Critères de rendement centre commercial communautaire à mail fermé||X||X"
		LabelArr(40) = "0|Frais de gestion - % RBE||X||"
		LabelArr(41) = "0|Critères de rendements - Édifice commercial à locataire unique|||X|"
		LabelArr(42) = "0|Motifs d'investissement sur le marché des centres commerciaux|X|||"
		LabelArr(43) = "0|Délai d'inoccupation - Excluant la période d'aménagement du commerce||||X"
		LabelArr(44) = "0|Délai d'inoccupation - Incluant la période d'aménagement du commerce||||X"
		LabelArr(45) = "0|Crédit accordé au commerçant||||X"
		LabelArr(46) = "0|Inflation centre commercial régional|X|||"
		LabelArr(47) = "0|Critères de rendement - Mégacentres|X||X|"
		LabelArr(48) = "0|Critères de rendement - Centre commercial de quartier avec épicerie|X||X|"
		LabelArr(49) = "0|Critères de rendement - Centre commercial régional de niveau 1|X|X|X|X"
		LabelArr(50) = "0|Critères de rendement - Centre commercial régional de niveau 2||X||X"
		LabelArr(51) = "0|Inflation - Centre commercial de niveau 2||||X"
		LabelArr(52) = "0|Acheteurs et vendeurs les plus actifs - Marché des centres commerciaux|||X|"
		LabelArr(53) = "0|Taux de retention des locataires commerciaux||||X"
		LabelArr(54) = "0|Inflation - Mégacentres|X|||"
		LabelArr(55) = "0|Inflation centre commercial communautaire à mail fermé||X||"
		LabelArr(56) = "0|Inflation - Centre commercial de quartier avec épicerie|X|||"
		LabelArr(57) = "0|Valeur à l'acquisition d'un terrain vacant pour développement d'un mégacentre|||X|"
		LabelArr(58) = "1|Industriel||||"
		LabelArr(59) = "0|Critères de rendement - Édifice industriel multi-locataires|X|X|X|X"
		LabelArr(60) = "0|Critères de rendement - Édifice industriel à locataire unique|X|X|X|X"
		LabelArr(61) = "0|Taux de location - Édifice industriel à locataire unique|X||X|"
		LabelArr(62) = "0|Inflation taux de location - Édifice industriel à locataire unique|X||X|"
		LabelArr(63) = "0|Taux de location initial - Édifice industriel multi-locataires||X||X"
		LabelArr(64) = "0|Inflation taux de location - Édifice industriel multi-locataires||X||X"
		LabelArr(65) = "0|Acheteurs et vendeurs les plus actifs - Marché industriel||||X"
		LabelArr(66) = "1|Édifices multifamiliaux||||"
		LabelArr(67) = "0|Critères de rendement - Édifice multifamilial|X|X|X|X"
		LabelArr(68) = "0|Dépenses d'exploitation pour la location - Édifice multifamilial|X|||"
		LabelArr(69) = "0|Construction multirésidentielle||X||"
		LabelArr(70) = "0|Rendement sur le développement et taux d'actualisation - Édifice multifamilial||X||"
		LabelArr(71) = "0|Inflation loyers - Édifice multifamilial|||X|"
		LabelArr(72) = "0|Critères de rendement - Édifice multifamilial au centre-ville|||X|"
		LabelArr(73) = "0|Édifice multifamilial en banlieue (coûts de roulement des locataires)||||X"
		LabelArr(74) = "0|Acheteurs/vendeurs - Édifice multifamilial de référence en tenure à bail|X|||"
		LabelArr(75) = "0|Critères de rendement - Édifice multifamilial en tenure à bail|X|||"
		LabelArr(76) = "0|Taux annuel de roulement des locataires||||X"
		LabelArr(77) = "0|Coûts de rénovation - Édifice multifamilial|X|||"
		LabelArr(78) = "0|Critères de rendement - Édifice multifamilial en périphérie du centre-ville de Montréal|X||X|"
		LabelArr(79) = "0|Facteurs de risque marché multifamilial|||X|"
		LabelArr(80) = "0|Motifs d'investissement sur le marché des édifices multifamiliaux||||X"
		LabelArr(81) = "1|Marché des capitaux||||"
		LabelArr(82) = "0|Sources de capital - Taux de rendement global exigé (FPI et sociétés cotées en bourse)||X||X"
		LabelArr(83) = "0|Disponibilité des capitaux propres sur le marché immobilier||X||X"
		LabelArr(84) = "0|Performance sur le prix du capital immobilier||X||X"
		LabelArr(85) = "0|Trois principales sources de capitaux pour prêts hypothécaires de premier rang|X||X|"
		LabelArr(86) = "0|Critères de marché - Emprunts conventionnels|X||X|"
		LabelArr(87) = "0|Perspectives sur le coût de la dette|X||X|"
		LabelArr(88) = "0|Provisions structurelles|X|||"
		LabelArr(89) = "0|Taux d'actualisation final à la revente|X|X||"
		LabelArr(90) = "0|Dépréciation recouvrable|||X|"
		LabelArr(91) = "0|Impact des taux d'intérêt sur les taux d'actualisation||X||"
		LabelArr(92) = "0|Déduire les frais de vente?||X||"
		LabelArr(93) = "0|Coûts de disposition à la revente||X||"
		LabelArr(94) = "1|Divers||||"
		   LabelArr(95) = "0|Intérêt partiel dans contrat de gestion excluant 20%, 33,3%, 50% d'intérêt|X|||"
		LabelArr(96) = "0|Indicateurs financiers pour édifices de classe AA au centre-ville||X||"
		LabelArr(97) = "0|Répartition terrain / valeur  - Critères de rendement d'il y a 10 ans||||X"
		LabelArr(98) = "0|Répartition terrain / valeur  - Critères de rendement d'il y a 30 ans||||X"
		LabelArr(99) = "0|Temps mise en marché - Nombre de mois|||X|"
		LabelArr(100) = "0|Profits sur honoraires de gestion||||X"
		LabelArr(101) = "0|Investissements prévus au Canada - Participation en capitaux et emprunts|X|||"
		LabelArr(102) = "0|Impact de l'homogénéité du portefeuille (escompte ou prime)||X||"
		   LabelArr(103) = "0|Valeur marchande estimée du portefeuille|X|||"


		Else


		ReDim LabelArr(100)
		LabelArr(0) = "1|Product Preferences||||"
		LabelArr(1) = "0|Altus InSite Location Barometer|X|X|X|X"
		LabelArr(2) = "0|Altus InSite Property Type Barometer|X|X|X|X"
		LabelArr(3) = "0|Altus InSite Product/Market Barometer|X|X|X|X"
		LabelArr(4) = "1|Office||||"
		LabelArr(5) = "0|Office Realty Commissions||X||"
		LabelArr(6) = "0|Office Tenant Retention Ratio|X|||"
		LabelArr(7) = "0|Office Lag Vacancy|||X|"
		LabelArr(8) = "0|Downtown Class ""AA"" Net Effective/Face Rates||X||X"
		LabelArr(9) = "0|Downtown Class ""AA"" Face Rate Inflation||X||X"
		LabelArr(10) = "0|Downtown Class ""B"" Office Valuation Parameters||X||"
		LabelArr(11) = "0|Downtown Land Parcel - Office Use||X||X"
		LabelArr(12) = "0|Leasehold Downtown Class ""AA"" Valuation Parameters|X|||"
		LabelArr(13) = "0|Downtown ""AA"" Economic Face Rate & Development Yield||||X"
		LabelArr(14) = "0|Suburban Class ""A"" Office Valuation Parameters|||X|"
		LabelArr(15) = "0|Suburban ""A"" Net Effective/Face Rates|||X|"
		LabelArr(16) = "0|Suburban ""A"" Face Rate Inflation|||X|"
		LabelArr(17) = "0|Suburban ""A"" Economic Face Rate & Development Yield|||X|"
		LabelArr(18) = "0|Suburban Office Land Parcel|X||X|"
		LabelArr(19) = "0|Downtown Class AA Office Valuation Parameters|X|X|X|X"
		LabelArr(20) = "0|Most Active Buyers and Sellers of Office Properties|||X|"
		LabelArr(21) = "0|Montreal Midtown Class ""A"" Office Valuation Parameters (MONTREAL ONLY)|X|X|X|X"
		LabelArr(22) = "0|Montreal Midtown Class ""A"" Office Face Rate Inflation (MONTREAL ONLY)|X|X|X|X"
		LabelArr(23) = "0|Suburban Class ""B"" Net Effective / Face Rates|X|||"
		LabelArr(24) = "0|Suburban Class ""B"" Face Rate Inflation|X|||"
		LabelArr(25) = "0|Office Inducement & Realty Commission Inflation - Downtown Class ""AA"" Office||X||"
		LabelArr(26) = "0|Suburban Class ""B"" Office Valuation Parameters|X|||"
		LabelArr(27) = "0|Third Party Property Management Fees (% of Effective gross income)||||X"
		LabelArr(28) = "0|Third Party Asset Management Fees (% of Asset Value)||||X"
		LabelArr(29) = "0|Toronto New Construction - Downtown Class ""AA"" & Suburban Class ""A"" Office Bldg|X|||"
		LabelArr(30) = "0|Motivation for Office Investment|||X|"
		LabelArr(31) = "0|Discount Rate to amortize tenant concessions, free rent or other forms of inducements||||X"
		LabelArr(32) = "0|Downtown Class ""B"" Net Effective/Face Rates||X||"
		LabelArr(33) = "0|Downtown Class ""B"" Face Rate Inflation||X||"
		LabelArr(34) = "0|Ground Lease Downtown Class ""AA"" Office Valuation Parameters||X||"
		LabelArr(35) = "0|Gross Rent to Sales Ratio|||X|"
		LabelArr(36) = "0|Insite-Altus Office Vacancy Barometer|X|X|X|X"
		LabelArr(37) = "1|Retail||||"
		LabelArr(38) = "0|Enclosed Community Mall Valuation Parameters||X||X"
		LabelArr(39) = "0|Retail Management Fees - % Eff Gross Inc||X||"
		LabelArr(40) = "0|Single-Tenant Retail Valuation Parameters|||X|"
		LabelArr(41) = "0|Motivation for Retail Investment|X|||"
		LabelArr(42) = "0|Retail Lag Vacancy - Excluding Fixturing Period||||X"
		LabelArr(43) = "0|Retail Lag Vacancy - Fixturing Period||||X"
		LabelArr(44) = "0|Retail Credit Allowance||||X"
		LabelArr(45) = "0|Regional Mall Inflation|X|||"
		LabelArr(46) = "0|Power Centre Valuation Parameters|X||X|"
		LabelArr(47) = "0|Food Anchored Retail Strip Valuation Parameters|X||X|"
		LabelArr(48) = "0|Tier I Regional Mall Valuation Parameters|X|X|X|X"
		LabelArr(49) = "0|Tier II Regional Mall Valuation Parameters||X||X"
		LabelArr(50) = "0|Tier II Regional Mall Inflation||||X"
		LabelArr(51) = "0|Most Active - Buyers & Sellers of Retail Properties|||X|"
		LabelArr(52) = "0|Retail Tenant Retention Ratio||||X"
		LabelArr(53) = "0|Power Centre Inflation|X|||"
		LabelArr(54) = "0|Enclosed Community Mall Inflation||X||"
		LabelArr(55) = "0|Food Anchored Retail Strip Inflation|X|||"
		LabelArr(56) = "0|Vacant Land Value for acquisition of similar Power Centre site|||X|"
		LabelArr(57) = "1|Industrial||||"
		LabelArr(58) = "0|Multi-Tenant Industrial Valuation Parameters|X|X|X|X"
		LabelArr(59) = "0|Single-Tenant Industrial Valuation Parameters|X|X|X|X"
		LabelArr(60) = "0|Rental Rate Single Tenant Industrial|X||X|"
		LabelArr(61) = "0|Rental Rate Inflation - Single Tenant Industrial|X||X|"
		LabelArr(62) = "0|Initial Net Rental Rate - Multi-Tenant Industrial||X||X"
		LabelArr(63) = "0|Rental Rate Inflation - Multi-Tenant Industrial||X||X"
		LabelArr(64) = "0|Most Active - Buyers & Sellers of Industrial Properties||||X"
		LabelArr(65) = "1|Multiple Unit Residential||||"
		LabelArr(66) = "0|Multi-Unit Residential Valuation Parameters|X|X|X|X"
		LabelArr(67) = "0|Multi-Unit Residential Rental Building Expenses|X|||"
		LabelArr(68) = "0|Multi-Residential Construction||X||"
		LabelArr(69) = "0|Multi-Residential Development Yield & Cap Rate||X||"
		LabelArr(70) = "0|Multi-Residential Rental Inflation|||X|"
		LabelArr(71) = "0|Downtown Multi-Unit Residential Valuation Parameters|||X|"
		LabelArr(72) = "0|Suburban Multi-Unti Residential (Suite Turnover Cost)||||X"
		LabelArr(73) = "0|Buyer/Seller of Benchmark Multi-Unit Residential Leasehold|X|||"
		LabelArr(74) = "0|Multi-Unit Residential Leasehold Parameters|X|||"
		LabelArr(75) = "0|Annual Tenant Turnover||||X"
		LabelArr(76) = "0|Multi-Unit Residential Rental - Renovation Cost|X|||"
		LabelArr(77) = "0|Midtown Multi-Unit Residential Valuation Parameters - Montreal|X||X|"
		LabelArr(78) = "0|Threats to Multi-Residential Investment|||X|"
		LabelArr(79) = "0|Motivation for Multi Residential Investment||||X"
		LabelArr(80) = "1|Capital Markets||||"
		LabelArr(81) = "0|Capital Sources - Total Required Return (REIT's & Publicly Traded Corp)||X||X"
		LabelArr(82) = "0|Real Estate Equity Available||X||X"
		LabelArr(83) = "0|Real Estate Equity Price Performance||X||X"
		LabelArr(84) = "0|Three Most Active sources of First Mortgage Debt Capital|X||X|"
		LabelArr(85) = "0|Conventional Debt Market Parameters|X||X|"
		LabelArr(86) = "0|Debt Cost Outlook|X||X|"
		LabelArr(87) = "0|Structural Reserve|X|||"
		LabelArr(88) = "0|Reversion Terminal Cap Rate|X|X||"
		LabelArr(89) = "0|Recoverable Depreciation|||X|"
		LabelArr(90) = "0|Interest Rate Impact on Capitalization Rates||X||"
		LabelArr(91) = "0|Deduct Selling Costs?||X||"
		LabelArr(92) = "0|Total Sales Cost at Reversion||X||"
		LabelArr(93) = "1|Miscellaneous||||"
		LabelArr(94) = "0|Financial Indicators for Downtown Class ""AA""||X||"
		LabelArr(95) = "0|Land / Valuation Parameters Split - 10 Year Old Valuation Parameters||||X"
		LabelArr(96) = "0|Land / Valuation Parameters Split - 30 Year Old Valuation Parameters||||X"
		LabelArr(97) = "0|Marketing Time - No. of Months|||X|"
		LabelArr(98) = "0|Management Fee Profits||||X"
		LabelArr(99) = "0|Planned Investment in Canada - Ownership & Debt|X|||"
		LabelArr(100) = "0|Homogeneous Portfolio Effect (Discount or Premium)||X||"

		End If

	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub


	Public Function Read()

		Read = TableObj.TableOpen(0)

		Read = Read & "<tr><th align=""left"" class=""head2"">" & Lang("topics") & "</th><th colspan=""4"" class=""head2"">" & Lang("qtr_frequency") & "</th></tr>"


		for i=0 to ubound(LabelArr)
			s=split(LabelArr(i), "|")

			if s(0)=1 then
				TableObj.RowID=""
				Read = Read & "<tr><th align=""left"" class=""head"">" & Ucase(s(1)) & "</th><th class=""head"">Q1</th><th class=""head"">Q2</th><th class=""head"">Q3</th><th class=""head"">Q4</th></tr>"
			else
				TableObj.RowID=s(1)
				Read = Read & TableObj.RowOpen & TableObj.Cell(s(2)) & TableObj.Cell(s(3)) & TableObj.Cell(s(4)) & TableObj.Cell(s(5)) & TableObj.RowClose
			end if
		next

		Read = Read & TableObj.TableClose

	End Function

End Class

%>
