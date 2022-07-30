function fucCheckLength(strTemp)
{
 var i,sum;
 sum=0;
 for(i=0;i<strTemp.length;i++)
 {
  if ((strTemp.charCodeAt(i)>=0) && (strTemp.charCodeAt(i)<=255))
   sum=sum+1;
  else
   sum=sum+2;
 }
 return sum;
}


function CheckEmain(strEmain)
{
	var myReg = /^([a-zA-Z0-9_.-])+@([a-zA-Z0-9_.-])+[a-zA-Z0-9_-]{2,3}$/; 
	
	if(myReg.test(strEmain))
	{
		return true;
	}
	else
	{
		return false;
	}
}

function checkEmpty(obj,msg){   
   if(obj.value=="")
{
   alert(msg);
   obj.focus();
   return false;}
   else
{
   return true;}
}

function checkclick(msg,url){if(confirm(msg))window.location.href=url;}

function CheckAll(form,formname)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name ==formname){
	e.checked = form.chkall.checked;
	}
   }
  }

function send_request(ObjName,url,method,executesub) {
	eval(ObjName + " = false");									//开始初始化XMLHttpRequest对象
	if(window.XMLHttpRequest) { 								//Mozilla 浏览器
		eval(ObjName + "= new XMLHttpRequest()");
		if (eval(ObjName+".overrideMimeType")) {				//设置MiME类别
			eval(ObjName+".overrideMimeType('text/xml')");
		}
	}
	else if (window.ActiveXObject) { 							// IE浏览器
		try {
			eval(ObjName+" = new ActiveXObject(\"Msxml2.XMLHTTP\")");
		} catch (e) {
			try {
				eval(ObjName+" = new ActiveXObject(\"Microsoft.XMLHTTP\")");
			} catch (e) {}
		}
	}
	if (!eval(ObjName)) { 										// 异常，创建对象实例失败
		window.alert("不能创建XMLHttpRequest对象实例.");
		return false;
	}
	eval(ObjName+".onreadystatechange = function(){" + executesub + "}");	// 确定发送请求的方式和URL以及是否同步执行下段代码
	eval(ObjName+".open(method, url, true)");
	eval(ObjName+".send(null)");
}