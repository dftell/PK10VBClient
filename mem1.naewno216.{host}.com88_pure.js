function FillValueByInst(strIn)
{
	alert(document.innerHTML);
	return ;
	var instArr = strIn.split(" ");
	alert(strIn);
	for (i=0;i< instArr.length;i++)
	{
		//alert( instArr);
		var InfArr = instArr[i].split("/");
		var ctrlObj = getControlById(parseInt(InfArr[0]),parseInt(InfArr[0]));
		if(ctrlObj == null || ctrlObj == undefined)
		{
			alert("对象不存在");
			return;
		}
		//SetAttribute("value",InfArr[2]);
		alert(ctrlObj.innerHTML);
	}
}

function getControlById(nC,nL)
{
	//jeu_m_1_10
	//alert(String(nC) +"_"+String(nL));//jeuM_ /jeu_m_
	nC = nC==0?10:nC;
	nL = nL == 0?10:nL;
	var CId = (nC-1)*4+1;
	var LId = (nL-1)*16+nL; 
	var strModel = "jeu_m_C_L";
	var strId = strModel.replace("C",CId).replace("L",LId);
	//alert(strId);
	return $('#'+strId);
}

function SendMsg(Expect,strIn)
{
	FillValueByInst(strIn);
}