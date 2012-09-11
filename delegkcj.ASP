<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title>新建网页 1</title>
<script language="javascript">
function change(xx)
{
if(xx.value=="1")
{pp.innerHTML="请输入姓名"
pp1.innerHTML="<input type='text' name='va' size='12' style='font-family: 华文行楷; font-size: 14pt'>"}

else if(xx.value=="2")
{pp.innerHTML="请输入准考证号"
pp1.innerHTML="<input type='text' name='va' size='12' style='font-family: 华文行楷; font-size: 14pt'>"}

else if(xx.value=="3")
{pp.innerHTML="请输入班级"
pp1.innerHTML="<select size='1' name='va' style='font-family: 宋体; font-size: 12pt' >"+
  "<option>05计一</option>"+
  "<option >05计二</option>"+
  "<option>05计三</option>"+
  "<option>05计四</option>"+
  "<option>04旅游</option>"+
  "<option>04机械</option>"+
  "<option>04服装</option>"+
  "</select>"
}
else if(xx.value=="4")
{pp.innerHTML="请输入报名号"
pp1.innerHTML="<input type='text' name='va' size='12' style='font-family: 华文行楷; font-size: 14pt'>"}



else if(xx.value=="5")
{pp.innerHTML="请输入专业"
pp1.innerHTML="<select size='1' name='va' style='font-family: 宋体; font-size: 12pt' >"+
  "<option>计算机</option>"+
    "<option>旅游</option>"+
  "<option>机械</option>"+
  "<option>服装</option>"+
  "</select>"
}


}


</script>

</head>

<body>

<p><font color="#FFFF00"><span style="background-color: #0033CC">选查询项目</span></font></p>
<form method="POST" action="delegkcj1.asp" target="right" >
  <select size="1" name="item" style="font-family: 宋体; font-size: 14pt" onChange="change(this)">
  <option selected value='1'>姓名</option>
  <option value='2'>准考证号</option>
  <option value='3'>班级</option>
  <option value='4'>报名号</option>
  <option value='5'>专业</option>
  </select></p>
<font color="#FF0000"><span style="background-color: #3399FF"><div id=pp>请输入姓名</div></span></font>

<p><font size="5">
<div id=pp1>
<input type="text" name="va" size="12" style="font-family: 华文行楷; font-size: 14pt">
</div>
</font></p>
<input type="submit" name="submit" value="执行删除命令">
</form>


</body>

</html>