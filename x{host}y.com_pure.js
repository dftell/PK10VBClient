function FillValueByInst(strIn)
{
	var instArr = strIn.split(" ");
	var i=0;
	for (i=0;i<instArr.length;i++)
	{
		var InfArr = instArr[i].split("/");
		var ctrlObj = getControlById(infArr[0],infArr[0]);
		ctrlObj.innerText =infArr[2];
	}
}

function getControlById(nC,nL)
{
	//jeu_m_1_10
	nC = nC==0?10:nC;
	nL = nL == 0?10:nL;
	var CId = (nC-1)*4+nC;
	var LId = (nL-1)*16+nL; 
	var strModel = "jeu_m_C_L";
	return Document.getElementById(strModel.replace("C",CId).replace("L",LId));
}

function SendMsg(Expect,strIn)
{
	FillValueByInst(strIn);
}