function Translate(strIn)
{
	var txt = eval('(' + strIn+ ')'); 
	var res = EnCryptJsonParams(txt);
	return 'info=' + res.info +  '&hasZip=' + res.hasZip;	
}
function TranslateInst(ExpertNo,__Token,strIn)
{
		//n=未知 8 行数
		//c=期号
		//totalTimesInput=整体投注倍数
		//l=具体组合
	    //gl.bet.betMode = 0 0，普通，1，倍数追号，2，高级追号
	    //e = [n, c, $totalTimesInput.val(), l, gl.bet.betMode, 1, "", ""];
	    	var n, c, l;
		n = "8";
		c = ExpertNo;
		l = eval('(' + strIn + ')'); 
		r = __Token;
		e=[n,c,"1",l,0,1,"",""];
		var res = EnCryptJsonParams(e);
		return '__RequestVerificationToken=' + r +  '&hasZip=' + res.hasZip +'&info=' + res.info  ;	
}

function Login(strUserName,strPassword)
{
    ctrlUsr =  document.getElementById("txt_username");
    ctrlPwd = document.getElementById("txt_pwd");
    ctrlbtn = document.getElementById("login-submit-button");
    ctrlUsr.value = strUserName;
    ctrlPwd.value = strPassword;
    ctrlbtn.click();
}

function SendMsg(strExpertNo,strIn)
{
	var n, c, l;
	n = "8";
	c = strExpertNo;
	l = eval('(' + strIn + ')'); 
	//r = __Token;
	e=[n,c,"1",l,0,1,"",""];
	//alert(EnCryptJsonParams(e).info);
	//var res = EnCryptJsonParams(e);
	//return '__RequestVerificationToken=' + r +  '&hasZip=' + res.hasZip +'&info=' + res.info  ;
	ctx.postTokenEx(
	{
		url:"/Bet/CqcSubmit",data:EnCryptJsonParams(e),
		beforeSend:function()
		{
			
		},
		complete:function()
		{
			
		},
		success:function(n)
		{
			//alert(n.Tip);
			n.Ok==1
			?(
				
				lv=n.gamePoint.toFixedNum(3),
				u="成功",
				t=3,
				gl.doLoop
				(
					{
						loopInt:t,
						backFunc:function()
						{
							$("#banIssueId").html(
							'状态：{0}，剩余：{1};'.replaceFormat([u,lv])
							),gl.baseBettingBanTips._show()
						}
					}
				),
				refreshGamePoint(n.gamePoint.toFixedNum(3))
			)
			:(
				//refreshGamePoint(n.gamePoint.toFixedNum(3)),
				i=n.Tip,
				u="失败",
				//$.appAlert({useTitle:"投注结果",title:u,message:i},1e3)
				t=3,
				gl.doLoop
				(
					{
						loopInt:t,
						backFunc:function()
						{
							$("#banIssueId").html(
							'状态：{0}，详细：{1}'.replaceFormat([i,u])
							),
							gl.baseBettingBanTips._show()
						}
					}
				)
			);
		}
			
	}
	)

}