<%

' DK, Aug 20, 2008 - version 1.0

Class cTableFrequency

	Private TableObj
	Private LabelArr

	' Constructor
	Private Sub Class_Initialize()

		Set TableObj=new cTableDatagrid
		TableObj.ReadOnly=True


		' RowType | Label| ChartTitle | Q1 | Q2 | Q3 | Q4

		If gLang="fre" Then

		ReDim LabelArr(85)
		LabelArr(0) =  "1||PRÉFÉRENCES PAR PRODUIT|Q1-09|Q2-09|Q3-09|Q4-09"
		LabelArr(1) =  "0|97APMB*|Baromètre Altus InSite par ville|X|X|X|X"
		LabelArr(2) =  "0|97APMB*|Baromètre Altus InSite par genre de propriété|X|X|X|X"
		LabelArr(3) =  "0|97APMB|Baromètre Altus InSite par produit/marché|X|X|X|X"
		LabelArr(4) =  "1||BUREAUX||||"
		LabelArr(5) =  "0|3VPO|Taux de rétention des locataires de bureaux|X|||"
		LabelArr(6) =  "0|4VPO|Délai d'inoccupation bureaux|||X|"
		LabelArr(7) =  "0|21VPOa|Loyers nets effectifs - Édifices de bureaux classe AA au centre-ville||X||X"
		LabelArr(8) =  "0|21VPOb*|Loyers nets demandés - Édifices de bureaux classe AA au centre-ville||X||X"
		LabelArr(9) =  "0|22VPO|Inflation loyers nets demandés - Édifices de bureaux classe AA au centre-ville||X||X"
		LabelArr(10) = "0|23VPO|Critères de rendement - Édifices de bureaux classe B au centre-ville||X||"
		LabelArr(11) = "0|25VPO|Terrain pour bureaux au centre-ville||X||X"
		LabelArr(12) = "0|55VPO|Critères de rendement - Édifice classe AA au centre-ville en tenure à bail|X|||"
		LabelArr(13) = "0|57VPOa|Loyer net économique - Édifices de bureaux classe AA au centre-ville||||X"
		LabelArr(14) = "0|57VPOb*|Rendement sur le développement - Édifices de bureaux classe AA au centre-ville||||X"
		LabelArr(15) = "0|64VPO|Critères de rendement - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(16) = "0|65VPOa|Loyers nets effectifs - Édifice de bureaux classe A en banlieue|||X|"
		LabelArr(17) = "0|65VPOb*|Loyers nets demandés - Édifice de bureaux classe A en banlieue|||X|"
		LabelArr(18) = "0|66VPO|Inflation loyers nets demandés - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(19) = "0|67VPOa|Loyer net économique - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(20) = "0|67VPOb*|Rendement sur le développement - Édifices de bureaux classe A en banlieue|||X|"
		LabelArr(21) = "0|99VPO|Critères de rendement - Édifices de bureaux classe AA au centre-ville|X|X|X|X"
		LabelArr(22) = "0|119VPO|Critères de rendement - Édifices classe A en périphérie du centre-ville de Montréal|X|X|X|X"
		LabelArr(23) = "0|120VPO|Loyers nets effectifs/loyers nets demandés - Édifices classe A en périphérie du centre-ville de Montréal|X|X|X|X"
		LabelArr(24) = "0|121VPO|Inflation loyers nets demandés - Édifices classe A en périphérie du centre-ville de Montréal|X|X|X|X"
		LabelArr(25) = "0|122VPOa|Loyers nets effectifs - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(26) = "0|122VPOb*|Loyers nets demandés - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(27) = "0|123VPO|Inflation loyers nets demandés - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(28) = "0|127VPO|Critères de rendement - Édifices de bureaux classe B en banlieue|X|||"
		LabelArr(29) = "0|128VPO|Honoraires de gestion par un tiers (% du revenu brut effectif)||||X"
		LabelArr(30) = "0|130VPO|Nouveaux projets à Toronto - Édifices classe AA au centre-ville et édifices classe B en banlieue|X|||"
		LabelArr(31) = "0|131VPO|Motifs d'investissement sur le marché des bureaux|||X|"
		LabelArr(32) = "0|132VPO|Fréquence d'escompte des loyers nets effectifs - Édifices de classe AA au centre-ville||||X"
		LabelArr(33) = "0|134VPO|Loyers nets effectifs/Loyers nets demandés - Édifices classe B au centre-ville||X||"
		LabelArr(34) = "0|135VPO|Inflation loyers nets demandés - Édifices de classe B au centre-ville||X||"
		LabelArr(35) = "0|136VPO|Critères de rendement - Bail à construction pour un édifice de bureaux classe AA au centre-ville||X||"
		LabelArr(36) = "0|103VPO|Baromètre Altus-InSite sur l'inoccupation de bureaux|X|X|X|X"
		LabelArr(37) = "1||CENTRES COMMERCIAUX||||"
		LabelArr(38) = "0|27VPR|Ratio ventes du loyer brut sur |||X|"
		LabelArr(39) = "0|28VPR|Critères de rendement centre commercial communautaire à mail fermé||X||X"
		LabelArr(40) = "0|30VPR|Frais de gestion - % RBE||X||"
		LabelArr(41) = "0|33VPR|Critères de rendements - Édifice commercial à locataire unique|||X|"
		LabelArr(42) = "0|34VPR|Motifs d'investissement sur le marché des centres commerciaux|X|||"
		LabelArr(43) = "0|59aVPR|Délai d'inoccupation - Incluant la période d'aménagement du commerce||||X"
		LabelArr(44) = "0|60VPR|Crédit accordé au commerçant||||X"
		LabelArr(45) = "0|61VPR|Inflation centre commercial régional|X|||"
		LabelArr(46) = "0|69VPR|Critères de rendement - Mégacentres|X||X|"
		LabelArr(47) = "0|70VPR|Critères de rendement - Centre commercial de quartier avec épicerie|X||X|"
		LabelArr(48) = "0|100VPR|Critères de rendement - Centre commercial régional de niveau 1|X|X|X|X"
		LabelArr(49) = "0|101VPR|Critères de rendement - Centre commercial régional de niveau 2||X||X"
		LabelArr(50) = "0|102VPR|Inflation - Centre commercial de niveau 2||||X"
		LabelArr(51) = "0|124aVPR|Taux de retention des locataires commerciaux||||X"
		LabelArr(52) = "0|125VPR|Inflation - Mégacentres|X|||"
		LabelArr(53) = "0|126VPR|Inflation centre commercial communautaire à mail fermé||X||"
		LabelArr(54) = "0|127VPR|Inflation - Centre commercial de quartier avec épicerie|X|||"
		LabelArr(55) = "0|139VPR|Valeur d'un terrain vacant pour développement d'un mégacentre|||X|"
		LabelArr(56) = "1||INDUSTRIEL||||"
		LabelArr(57) = "0|33VPI|Critères de rendement - Édifice industriel multi-locataires|X|X|X|X"
		LabelArr(58) = "0|75VPI|Critères de rendement - Édifice industriel à locataire unique|X|X|X|X"
		LabelArr(59) = "0|128VPI|Taux de location - Édifice industriel à locataire unique|X||X|"
		LabelArr(60) = "0|129VPI|Inflation taux de location - Édifice industriel à locataire unique|X||X|"
		LabelArr(61) = "0|133VPI|Taux de location initial - Édifice industriel multi-locataires||X||X"
		LabelArr(62) = "0|134VPI|Inflation taux de location - Édifice industriel multi-locataires||X||X"
		LabelArr(63) = "1||ÉDIFICES MULTIFAMILIAUX||||"
		LabelArr(64) = "0|36VPMR|Critères de rendement - Édifice multifamilial en banlieue|X|X|X|X"
		LabelArr(65) = "0|38VPMR|Dépenses d'exploitation pour la location - Édifice multifamilial en banlieue|X|||"
		LabelArr(66) = "0|40VPMR|Rendement sur le développement et taux d'actualisation - Édifice multifamilial||X||"
		LabelArr(67) = "0|81VPMR|Critères de rendement - Édifice multifamilial au centre-ville|||X|"
		LabelArr(68) = "0|82VPMR|Édifice multifamilial en banlieue (coûts de roulement des locataires)||||X"
		LabelArr(69) = "0|117VPMR|Critères de rendement - Édifice multifamilial en tenure à bail|X|||"
		LabelArr(70) = "0|118VPMR|Taux annuel de roulement des locataires||||X"
		LabelArr(71) = "0|127VPMR|Critères de rendement - Édifice multifamilial en périphérie du centre-ville de Montréal|X||X|"
		LabelArr(72) = "0|128VPMR|Facteurs de risque marché multifamilial|||X|"
		LabelArr(73) = "0|129VPMR|Motifs d'investissement sur le marché des édifices multifamiliaux||||X"
		LabelArr(74) = "1||MARCHÉ DES CAPITAUX||||"
		LabelArr(75) = "0|43CSCM|Taux de rendement global exigé (FPI et sociétés cotées en bourse)||X||X"
		LabelArr(76) = "0|47CSCM|Disponibilité des capitaux propres pour sociétés immobilières de toute grandeur||X||X"
		LabelArr(77) = "0|48CSCM|Performance sur le prix du capital immobilier||X||X"
		'LabelArr() = "0|82CSDM|Trois principales sources de capitaux pour prêts hypothécaires de premier rang|X||X|"
		LabelArr(78) = "0|83CSDM|Critères de marché - Emprunts conventionnels|X||X|"
		LabelArr(79) = "0|85aCSDM|Perspectives sur le coût de la dette|X||X|"
		LabelArr(80) = "1||DIVERS||||"
		LabelArr(81) = "0|18MISC|Impact des taux d'intérêt sur les taux d'actualisation||X||"
		'LabelArr() = "0|49MISC|Calcul de la valeur de réversion||X||"
		'LabelArr() = "0|50MISC|Coûts de disposition à la réversion||X||"
		LabelArr(82) = "0|89MISC|Indicateurs financiers pour édifices de classe AA au centre-ville||X||"
		LabelArr(83) = "0|96MISC|Temps mise en marché - Nombre de mois|||X|"
		LabelArr(84) = "0|108MISC|Investissements prévus au Canada - Participation en capitaux et emprunts|X|||"
		LabelArr(85) = "0|110MISC|Impact de l'homogénéité du portefeuille (escompte ou prime)||X||"


		Else


		ReDim LabelArr(85)
		LabelArr(0) =  "1||PRODUCT PREFENCES||||"
		LabelArr(1) =  "0|97APMB*|Altus InSite Market Barometer|X|X|X|X"
		LabelArr(2) =  "0|97APMB*|Altus InSite Product Barometer|X|X|X|X"
		LabelArr(3) =  "0|97APMB|Altus InSite Product/Market Barometer|X|X|X|X"
		LabelArr(4) =  "1||OFFICE||||"
		LabelArr(5) =  "0|3VPO|Office Tenant Retention Ratio|X|||"
		LabelArr(6) =  "0|4VPO|Office Lag Vacancy|||X|"
		LabelArr(7) =  "0|21VPOa|Downtown Class ""AA"" Office: Net Effective Rates||X||X"
		LabelArr(8) =  "0|21VPOb*|Downtown Class ""AA"" Ofiice: Market Face Rates||X||X"
		LabelArr(9) =  "0|22VPO|Downtown Class ""AA"" Office: Face Rate Inflation||X||X"
		LabelArr(10) = "0|23VPO|Downtown Class ""B"" Office Valuation Parameters||X||"
		LabelArr(11) = "0|25VPO|Downtown Land Parcel - Office Use||X||X"
		LabelArr(12) = "0|55VPO|Leasehold Downtown Class ""AA"" Office|X|||"
		LabelArr(13) = "0|57VPOa|Downtown ""AA"" Office: Economic Face Rates||||X"
		LabelArr(14) = "0|57VPOb*|Downtown ""AA"" Office: Development Yields||||X"
		LabelArr(15) = "0|64VPO|Suburban Class ""A"" Office Valuation Parameters|||X|"
		LabelArr(16) = "0|65VPOa|Suburban Class ""A"" Office: Net Effective Rates|||X|"
		LabelArr(17) = "0|65VPOb*|Suburban Class ""A"" Office: Market Face Rates|||X|"
		LabelArr(18) = "0|66VPO|Suburban Class ""A"" Office: Face Rate Inflation|||X|"
		LabelArr(19) = "0|67VPOa|Suburban Class ""A"" Office: Economic Face Rates|||X|"
		LabelArr(20) = "0|67VPOb*|Suburban Class ""A"" Office: Development Yields|||X|"
		LabelArr(21) = "0|99VPO|Downtown Class AA Office Valuation Parameters|X|X|X|X"
		LabelArr(22) = "0|119VPO|Montreal Midtown Class ""A"" Valuation Parameters|X|X|X|X"
		LabelArr(23) = "0|120VPO|Montreal Midtown Class ""A"" Net Effective & Face Rates|X|X|X|X"
		LabelArr(24) = "0|121VPO|Montreal Midtown Class ""A"" Office Face Rate Inflation|X|X|X|X"
		LabelArr(25) = "0|122VPOa|Suburban Class ""B"" Office: Net Effective Rates|X|||"
		LabelArr(26) = "0|122VPOb*|Suburban Class ""B"" Office: Market Face Rates|X|||"
		LabelArr(27) = "0|123VPO|Suburban Class ""B"" Office Face Rate Inflation|X|||"
		LabelArr(28) = "0|127VPO|Suburban Class ""B"" Office Valuation Parameters |X|||"
		LabelArr(29) = "0|128VPO|Third Party Property Management Fees (% of Effective gross income)||||X"
		LabelArr(30) = "0|130VPO|Toronto New Construction - Downtown Class ""AA"" & Suburban Class ""A"" Office|X|||"
		LabelArr(31) = "0|131VPO|Motivation for Office Investment|||X|"
		LabelArr(32) = "0|132VPO|Downtown Class ""AA"" Office Net Effective Discount Period||||X"
		LabelArr(33) = "0|134VPO|Downtown Class ""B"" Office: Net Effective/Face Rates||X||"
		LabelArr(34) = "0|135VPO|Downtown Class ""B"" Office: Face Rate Inflation||X||"
		LabelArr(35) = "0|136VPO|Ground Lease Downtown Class ""AA"" Office Valuation Parameters||X||"
		LabelArr(36) = "0|103VPO|Office Vacancy Barometer|X|X|X|X"
		LabelArr(37) = "1||RETAIL||||"
		LabelArr(38) = "0|27VPR|Gross Rent to Sales Ratio |||X|"
		LabelArr(39) = "0|28VPR|Enclosed Community Mall Valuation Parameters||X||X"
		LabelArr(40) = "0|30VPR|Retail Property Third Party Management Fees||X||"
		LabelArr(41) = "0|33VPR|Single-Tenant Retail Valuation Parameters|||X|"
		LabelArr(42) = "0|34VPR|Motivation for Retail Investment|X|||"
		LabelArr(43) = "0|59aVPR|Retail Lag Vacancy - Including Fixturing Period||||X"
		LabelArr(44) = "0|60VPR|Retail Vacancy and Credit Allowance||||X"
		LabelArr(45) = "0|61VPR|Tier I Regional Mall Rent, Expense &Sales Inflation|X|||"
		LabelArr(46) = "0|69VPR|Power Centre Valuation Parameters|X||X|"
		LabelArr(47) = "0|70VPR|Food Anchored Retail Strip Valuation Parameters|X||X|"
		LabelArr(48) = "0|100VPR|Tier I Regional Mall Valuation Parameters|X|X|X|X"
		LabelArr(49) = "0|101VPR|Tier II Regional Mall Valuation Parameters (by province)||X||X"
		LabelArr(50) = "0|102VPR|Tier II Regional Mall Rent & Expense Inflation||||X"
		LabelArr(51) = "0|124aVPR|Retail Tenant Retention Ratio||||X"
		LabelArr(52) = "0|125VPR|Power Centre Rent & Expense Inflation|X|||"
		LabelArr(53) = "0|126VPR|Enclosed Community Mall Rent & Expense Inflation||X||"
		LabelArr(54) = "0|127VPR|Food Anchored Retail Strip Rent & Expense Inflation|X|||"
		LabelArr(55) = "0|139VPR|Vacant Land Value for Power Centre site|||X|"
		LabelArr(56) = "1||INDUSTRIAL||||"
		LabelArr(57) = "0|33VPI|Multi-Tenant Industrial Valuation Parameters|X|X|X|X"
		LabelArr(58) = "0|75VPI|Single Tenant Industrial Valuation Parameters|X|X|X|X"
		LabelArr(59) = "0|128VPI|Single Tenant Industrial Net Rental Face Rate|X||X|"
		LabelArr(60) = "0|129VPI|Single Tenant Industrial Rental Face Rate Inflation|X||X|"
		LabelArr(61) = "0|133VPI|Multi-Tenant Industrial Initial Net Rental Face Rate||X||X"
		LabelArr(62) = "0|134VPI|Multi-Tenant Industrial Rental Face Rate Inflation||X||X"
		LabelArr(63) = "1||MULTIPLE UNIT RESIDENTIAL||||"
		LabelArr(64) = "0|36VPMR|Suburban Multi-Unit Residential Valuation Parameters |X|X|X|X"
		LabelArr(65) = "0|38VPMR|Suburban Multi-Unit Residential Rental Building Expenses|X|||"
		LabelArr(66) = "0|40VPMR|Multi-Residential Development Yield & Cap Rate||X||"
		LabelArr(67) = "0|81VPMR|Downtown Multi-Unit Residential Valuation Parameters |||X|"
		LabelArr(68) = "0|82VPMR|Suburban Multi-Unti Residential (Suite Turnover Cost)||||X"
		LabelArr(69) = "0|117VPMR|Suburban Multi-Unit Residential Leasehold Parameters |X|||"
		LabelArr(70) = "0|118VPMR|Suburban Multi-Unit Residential: Annual Tenant Turnover ||||X"
		LabelArr(71) = "0|127VPMR|Montreal Midtown Multi-Unit Residential Building Valuation Parameters|X||X|"
		LabelArr(72) = "0|128VPMR|Threats to Multi-Unit Residential Investment|||X|"
		LabelArr(73) = "0|129VPMR|Motivation for Multi Residential Investment||||X"
		LabelArr(74) = "1||CAPITAL MARKETS||||"
		LabelArr(75) = "0|43CSCM|Total Required Return (REITs & Corporations)||X||X"
		LabelArr(76) = "0|47CSCM|New Real Estate Equity for large and small cap RE companies ||X||X"
		LabelArr(77) = "0|48CSCM|Expected Performance of Real Estate Equities ||X||X"
		'LabelArr() = "0|82CSDM|Three Most Active sources of First Mortgage Debt Capital|X||X|"
		LabelArr(78) = "0|83CSDM|Conventional Debt Market Parameters |X||X|"
		LabelArr(79) = "0|85aCSDM|Debt Cost Outlook|X||X|"
		LabelArr(80) = "1||MISCELLANEOUS||||"
		LabelArr(81) = "0|18MISC|Interest Rate Impact on Capitalization Rates||X||"
		'LabelArr() = "0|49MISC|Calculation of Reversion Value||X||"
		'LabelArr() = "0|50MISC|Selling Costs at Reversion||X||"
		LabelArr(82) = "0|89MISC|Downtown Class ""AA"" Office Financial Indicators ||X||"
		LabelArr(83) = "0|96MISC|Marketing Time - Number of Months|||X|"
		LabelArr(84) = "0|108MISC|Planned Investment in Canada - Ownership & Debt|X|||"
		LabelArr(85) = "0|110MISC|Homogeneous Portfolio Effect (Discount or Premium)||X||"

		End If

	End Sub

	' Destructor
	Private Sub Class_Terminate()
		Set TableObj=nothing
	End Sub


	Public Function Read()

		Read = TableObj.TableOpen(0)

		Read = Read & "<tr><th style=""text-align:left;"" class=""head2"">Library&nbsp;No.</th><th style=""text-align:left;"" class=""head2"">" & Lang("topics") & "</th><th colspan=""4"" class=""head2"">" & Lang("qtr_frequency") & "</th></tr>"


		for i=0 to ubound(LabelArr)
			s=split(LabelArr(i), "|")

			TableObj.RowID=""

			if s(0)=1 then
				Read = Read & "<tr><th style=""text-align:left;"" class=""head""></th><th style=""text-align:left;"" class=""head"">" & Server.HTMLEncode(Ucase(s(2))) & "</th><th class=""head"">Q1</th><th class=""head"">Q2</th><th class=""head"">Q3</th><th class=""head"">Q4</th></tr>"
			else
				Read = Read & "<tr align=""center""><th style=""text-align:left;"">" & Server.HTMLEncode(s(1)) & "</th><th style=""text-align:left;"">" & Server.HTMLEncode(s(2)) & "</th><td>" & Server.HTMLEncode(s(3)) & "</td><td>" & Server.HTMLEncode(s(4)) & "</td><td>" & Server.HTMLEncode(s(5)) & "</td><td>" & Server.HTMLEncode(s(6)) & "</td></tr>"
			end if
		next

		Read = Read & TableObj.TableClose

	End Function

End Class

%>
