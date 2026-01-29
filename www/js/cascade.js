/*
 * cascade.js
 *
 * Gets values of select box depending on previous selection
 *
 * depends on the Prototype javascript library. Tested with versions 1.5.1 and 1.6
 *
 */


function removeOptions(destObj)
{
	// Does not remove the first Option (the zeroeth Option)
	// for (var i=destObj.options.length-1; i>0; i--) destObj.remove(i);

	// Does remove the first Option (the zeroeth Option)
	for (var i=destObj.options.length-1; i>=0; i--) destObj.remove(i);
}

function addToSelect(destObj, newVals)
{
	for (var i in newVals) destObj.options[destObj.options.length] = new Option(newVals[i], i, false, false);
}


function getNextValues(objName, firstId, secondId, periodId, lang)
{
	url='/eng/chart/cascade.asp';
	if (lang=='fr-ca') url='/fre/chart/cascade.asp';
	
	new Ajax.Request(url,
	{
		method: 'get',

		parameters: {change: objName, firstid: firstId, secondid: secondId, period: periodId},

		onSuccess:
		
			function(transport, json)
			{
				if ($(objName).disabled)
				{
					$(objName).disabled=false;
				}
				else
				{
					removeOptions($(objName));
				};

				addToSelect($(objName), json);
			},
		
		onFailure:

			function()
			{
				alert('Data not returned.')
			}
	}
	);
}


function updateStep1(periodId, groupId, lang)
{
	if (groupId==1)
	{
		getNextValues('property_1', $F('report_1'), $F('product_1'), periodId, lang);
		getNextValues('market_1', $F('report_1'), $F('product_1'), periodId, lang);
	}
	else
	{
		getNextValues('property_0', $F('report_0'), $F('product_0'), periodId, lang);
		getNextValues('market_0', $F('report_0'), $F('product_0'), periodId, lang);
	}
	
	//updateMarket();
}

function updateMarket()
{
	destObj=$('market_1');

	var i, j=8, n=15, hist_investor_outlook_id='H100', hist_valuation_params_id='H101', tier_ii_regional_mall_id=8;

	if (($F('report_1')!=hist_investor_outlook_id) && ($F('report_1')!=hist_valuation_params_id)) return;

	if (destObj.options.length != n) return;

	if ($F('property_1')==tier_ii_regional_mall_id)
	{
		for (i=0; i<j; i++) destObj.options[i].disabled=true;
		for (i=j; i<n; i++) destObj.options[i].disabled=false;

		//for (i=0; i<j; i++) destObj.options[i].hide();
		//for (i=j; i<n; i++) destObj.options[i].show();

		destObj.selectedIndex=j;
	}
	else
	{
		for (i=0; i<j; i++) destObj.options[i].disabled=false;
		for (i=j; i<n; i++) destObj.options[i].disabled=true;

		//for (i=0; i<j; i++) destObj.options[i].show();
		//for (i=j; i<n; i++) destObj.options[i].hide();

		destObj.selectedIndex=0;
	}
}

function handleForm(my_form)
{
	my_form.action="result.asp";
	my_form.step.value="2";
	return true;
}

setTimeout(function(){
	updateMarket();
}, 2000);