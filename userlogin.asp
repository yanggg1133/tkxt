<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title>�½���ҳ 1</title>
<script language="javascript">
var i=0
function disp()
{if(i==0)
 {document.all.xs.style.display="block";i=1;}
else
{document.all.xs.style.display="none";i=0;}
}
function change(xx)
{
if(xx.value=="1")
{pp.innerHTML="����������"
pp1.innerHTML="<input type='text' name='va' size='12' style='font-family: �����п�; font-size: 14pt'>"}

else if(xx.value=="2")
{pp.innerHTML="������׼��֤��"
pp1.innerHTML="<input type='text' name='va' size='12' style='font-family: �����п�; font-size: 14pt'>"}

else if(xx.value=="3")
{pp.innerHTML="������༶"
pp1.innerHTML="<select size='1' name='va' style='font-family: ����; font-size: 12pt' >"+
  "<option>05��һ</option>"+
  "<option >05�ƶ�</option>"+
  "<option>05����</option>"+
  "<option>05����</option>"+
  "<option>04����</option>"+
  "<option>04��е</option>"+
  "<option>04��װ</option>"+
  "</select>"
}
else if(xx.value=="4")
{pp.innerHTML="�����뱨����"
pp1.innerHTML="<input type='text' name='va' size='12' style='font-family: �����п�; font-size: 14pt'>"}



else if(xx.value=="5")
{pp.innerHTML="������רҵ"
pp1.innerHTML="<select size='1' name='va' style='font-family: ����; font-size: 12pt' >"+
  "<option>�����</option>"+
    "<option>����</option>"+
  "<option>��е</option>"+
  "<option>��װ</option>"+
  "</select>"
}


}


</script>

</head>

<body>

<p><font color="#FFFF00"><span style="background-color: #0033CC">ѡ��ѯ��Ŀ</span></font></p>
<form method="POST" action="cjdisp.asp" target="right" >
  <select size="1" name="item" style="font-family: ����; font-size: 14pt" onChange="change(this)">
  <option selected value='1'>����</option>
  <option value='2'>׼��֤��</option>
  <option value='3'>�༶</option>
  <option value='4'>������</option>
  <option value='5'>רҵ</option>
  </select></p>
<font color="#FF0000"><span style="background-color: #3399FF"><div id=pp>����������</div></span></font>

<p><font size="5">
<div id=pp1>
<input type="text" name="va" size="12" style="font-family: �����п�; font-size: 14pt">
</div>
</font></p>
<input type="submit" name="submit" value="��ѯ">
</form>

<form method="POST" action="admin.asp" target="left" >
<br>
<p>
<font color="#660033"><span style="background-color: #FFFF00" onclick="disp()">����Աע��</span></font>

<div id=xs style="display:none"> 
<br>�û���:<input type="text" name="user" size="6" style="font-family: �����п�; font-size: 11pt"></font>


<br>��&nbsp;&nbsp;��:<input type="password" name="pass" size="6" style="font-family: �����п�; font-size: 11pt"></font>
  
<br><input type="submit" name="submit" value="�ύ">
</div>
</form>

</body>

</html>