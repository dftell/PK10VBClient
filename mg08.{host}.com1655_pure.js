function Login(strUser,strPwd)
{
	$("#loginName").val(strUser);
	$("#loginPwd").val(strPwd);
}
function FillValueByInst(strIn)
{
	strIn = strIn.replace(/(^\s*)(\s*$)/g, "");
	var instArr = strIn.split(" ");
	//alert(strIn);
	var i=0;
	//alert(instArr);
	for (i=0;i<instArr.length;i++)
	{
		
		var strComb = instArr[i];
		
		var infArr = strComb.split("/");
		//var ctrlObj = 
		//alert(infArr);
		getControlById(infArr [0],infArr[1],infArr[2]);
		//alert(ctrlObj.innerText);
		//ctrlObj.innerText =infArr[2];
	}
	
}

function getControlById(nC,nL,val)
{
	//jeu_m_1_10
	nC = nC==0?10:nC;
	nL = nL == 0?10:nL;
	var CId = (nC-1)*4+nC;
	var LId = (nL-1)*16+nL; 
	var strModel = "#jeu_m_C_L";
	var strObj = strModel.replace("C",CId).replace("L",LId);
	strModel = "t{0}_h{1}";
	strObj = strModel.replace("{0}",nC).replace("{1}",nL);
	
	inobj = document.getElementsByName(strObj);
	//alert(strObj);
	//alert(inobj);
	if(inobj != null)
	{
		//alert("对象存在");
		//alert(inobj.type);
		//inobj.click();
		
		//Shortcut_ImportM(inobj);
		$(inobj).val(val);
		//digitOnly(inobj);
		//inobj.keyup();
	}
	else
	{
		//alert(inobj);
	}
	//return Document.getElementById(strModel.replace("C",CId).replace("L",LId));
}

function SendMsg(Expect,strIn)
{
	//alert(Expect + strIn);
	FillValueByInst(strIn);
	subobj = document.getElementById("submits");
	submitforms1(Expect);
}

function submitforms1(vExpect)
{
	
	$.post("../AjaxAll/Default.ajax_pk3.php", { typeid : "sessionId"}, function(){});
	
	var mixmoney = parseInt($("#mix").val()); //最低下注金額
	
	var input = $("input.inp1");
	
	var c = true, s, n;
	
	var count = 0;
	
	var countmoney = 0;

	var upmoney = 0;
	
	var names = new Array();
	
	var sArray = "";
	
	input.each(function()
		{
		
			var value = $(this).val();
		
			if (value != "")
			{
			
				value = parseInt(value);
			
				if (value < mixmoney) 
					c=false;
			
				count++;
			
				countmoney += value;
			
				s = nameformat($(this).attr("name").split("_"));
			
				s[2] = $("#"+s[2]+" a").html();
			
				/*if (s[0] == "總和、龍虎")
				
					n = s[1]+" @ "+s[2]+" x ￥"+value;
			
				else 
				
				n = s[0]+"["+s[1]+"] @ "+s[2]+" x ￥"+value;*/
			
			
				if (s[0] == "總和、龍虎")
						
					n = "<tr><td class='Ball_tr_H' width='30%'>"+"總和、龍虎"+"</td><td class='Ball_tr_H' width='40%'><span class='Font_B'>【"+s[1]+
"】</span> @<b class='Font_R'>"+s[2]+"</b></td><td class='Ball_tr_H Font_R' width='30%'>￥"+value+"</td></tr>";
			
				else 
							
					n = "<tr><td class='Ball_tr_H' width='30%'>"+s[0]+"</td><td class='Ball_tr_H' width='40%'><span class='Font_B'>【"+s[1]+
"】</span> @<b class='Font_R'>"+s[2]+"</b></td><td class='Ball_tr_H Font_R' width='30%'>￥"+value+"</td></tr>";
			
			
				names.push(n+"\n");
			
				sArray += s+","+value+"|";
		
			}
	
		});
	
	if (count == 0)
	{
		 
			
		var timer;
			
		var con = '請填寫下註金額!!!';

			
		if(con != '')
		{
				
			art.dialog({
					
				title: '警告',
					
				content: '<span style="font-size:15px">'+ con +'</span>',
					
				init: function ()
				 {
						
					var that = this, i = 10;
						
					var fn = function () {
							
					that.title('警告( ' + i + ' 秒后关闭)');
							
					!i && that.close();
							
					i --;
										
				};
						
				timer = setInterval(fn, 1000);
						
				fn();
					
			},
					
			close: function () {

				clearInterval(timer);

			},
					
			ok: function () {}
				
			}).show();

		}
 
		
	return;

	}
	
	
	
	if (c == false)
	{
		var timer;
	
		var con = "最低下註金額："+mixmoney+"￥";
		if(con != '')
		{
			art.dialog({

			title: '警告',

			content: '<span style="font-size:15px">'+ con +'</span>',

			init: function () 
			{

				var that = this, i = 10;

				var fn = function () 
				{
	
					that.title('警告( ' + i + ' 秒后关闭)');

					!i && that.close();

					i --;

				};

				timer = setInterval(fn, 1000);

				fn();

			},

			close: function () {
clearInterval(timer);
},

			ok: function () {
}

			}).show();

		}
 

 		return;
	}

	var numb = $("#o").html();
	numb = parseInt(numb);

	vExpect = parseInt(vExpect );
	if(numb<=0 || isNaN(numb)||numb!=vExpect)
	{

		var timer;
	
		var con = "期数获取失败，请重新下注！";
		if(con != '')
		{

			art.dialog(
			{

				title: '警告',

				content: '<span style="font-size:15px">'+ con +'</span>',

				init: function ()
				 {

					var that = this, i = 10;

					var fn = function () 
					{
	
						that.title('警告( ' + i + ' 秒后关闭)');

						!i && that.close();

						i --;

					};

					timer = setInterval(fn, 1000);

					fn();

				},

				close: function () {
clearInterval(timer);
},

				ok: function () {}

			}).show();

		}
 

		return
	}		
	
	var topstr = "<div class='ddmBox'><table class='alertTable Ball_List' width='100%'><tr><td class='td_caption_1' width='30%'>類型</td><td class='td_caption_1' width='40%'>明細</td><td class='td_caption_1' width='30%'>金額</td></tr></table><div class='ddBox'><table class='Ball_List' width='100%'>";
	  
	
	var bottomstr = "</table></div><div class='tjBox'>"+"合計：<b class='Font_R'>"+count+"</b>筆 共 ￥<b class='Font_R'>"+countmoney+"</b></div></div>";

		confrims  = topstr + names.join('') + bottomstr;
		
	input.val("");
			
	var number = $("#o").html();
			
	var s_type = '<input type="hidden" name="sm_arr" value="'+sArray+'"><input type="hidden" name="s_number" value="'+number+'">';
					$(".actiionn").html(s_type);
			
	$("#dp").submit();
					
	myLockShow();
			
	setTimeout(function () 
	{

		myLockHide('1')
			
	}, 1000);		 
		
}	
	

