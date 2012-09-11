<!--#INCLUDE FILE="conn.asp" -->

<%@language=vbscript 
%>
<script language="javascript">
function check(myform)
{if(myform.xm.value=="")
{alert("请输入姓名！");
myform.xm.focus();
return false;
}
 if(myform.xh.value=="")
{ alert("请输入学号！");
myform.xh.focus();
return false;
}
if(myform.bj.value=="")
{ alert("请输入班级！");
myform.bj.focus();
return false;
}
return true;
}
</script>
<% 
if application("exec")="" then

set rs1=server.createobject("adodb.recordset")
rs1.open "select * from sj where(status=1)" ,conn,3,3                                 

if not rs1.eof then                 
application("exec")=rs1("tx")   '考试试题                          
application("dh")=rs1("dh")      '试卷代号                       
application("sm")=rs1("sm")      '试卷标题                       
application("kssj")=rs1("kssj")   '考试时间                         
application("kshm")=rs1("kshm")  '考试号码                     
 end if  
 rs1.close
end if


if application("exec")="" then%>
<%response.expires=0
 response.clear%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="refresh" content="6,url=stuanswer.asp">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>das</title>
</head>
<body bgcolor="#F2FFE6">
<center>老师还没发试卷，请耐心等待.....
<br><br><p> <a href=kh.asp>-----返回-----</a></p></center>
</body>
</html>
<%else%>
<html>
<body>
<head>
</head>
<p></p>
<p></p>
<p></p>
<div align="center">
  <center>
  <form name="form1" method="post" action="stuanswer1.asp" onsubmit="return check(this)">
<table border="1" width="43%" height="119">
  <tr>
    <td width="100%" colspan="2" bgcolor="#0000FF" height="22">
      <p align="center"><font color="#FFFFFF">学生考试信息</font></td>
  </tr>
  <tr>
    <td width="43%" align="right" height="23">请输入<font color="#FF0000">姓名</font>:</td>
    <td width="57%" height="23"><input type ="text" name="xm" size="20"></td>
  </tr>
  <tr>
    <td width="43%" align="right" height="23">请输入<font color="#FF0000">学号</font>:</td>
    <td width="57%" height="23"><input type ="text" name="xh" size="20"></td>
  </tr>
  <tr>
    <td width="43%" align="right" height="15">请输入<font color="#FF0000">班级</font>:</td>
    <td width="57%" height="15"><input type ="text" name="bj" size="20"></td>
  </tr>
  <tr>
    <td width="43%" height="6"></td>
    <td width="57%" height="6"><input type="submit" name="submit" value="开始答题">
</td>
  </tr>
</table>
</form>
  </center>
</div>
</body>
</html>
<%end if%>