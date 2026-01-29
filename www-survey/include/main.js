function InputCallback(qID)
{
	if (qID.indexOf('MI108')>=0)
	{
		var ti=new Array("A", "B", "C", "D", "E", "F")
		var col = new Array("1O", "2O", "1D", "2D");
		var total = new Array(0, 0, 0, 0);
		
		for (i=0 ; i<6 ; i++)
		{
			for (j=0 ; j<4 ; j++)
			{
				inputbox =  document.getElementById('MI108' + ti[i] + col[j]);
				num = parseInt(inputbox.value);
				if(isNaN(num))
					num = 0;
				total[j] += num;
			}
		}

		var el10 = document.getElementById('Total1O');
		var el20 = document.getElementById('Total2O');
		var el1D = document.getElementById('Total1D');
		var el2D = document.getElementById('Total2D');
			
		el10.innerHTML  = total[0];
		el20.innerHTML  = total[1];
		el1D.innerHTML  = total[2];
		el2D.innerHTML  = total[3];
	}
}

function setCol(value, code)
{
	l=document.mainform.elements.length;

	for (i=0; i<l; i++)
	{
		if (document.mainform.elements[i].type != "text") continue;
		if (document.mainform.elements[i].name.indexOf(code)<0) continue;
		document.mainform.elements[i].value = value;
	}
}

