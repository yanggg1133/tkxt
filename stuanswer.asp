<!--#INCLUDE FILE="conn.asp" -->

<%@language=vbscript 
%>
<script language="javascript">
function check(myform)
{if(myform.xm.value=="")
{alert("������������");
myform.xm.focus();
return false;
}
 if(myform.xh.value=="")
{ alert("������ѧ�ţ�");
myform.xh.focus();
return false;
}
if(myform.bj.value=="")
{ alert("������༶��");
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
application("exec")=rs1("tx")   '��������                          
application("dh")=rs1("dh")      '�Ծ����                       
application("sm")=rs1("sm")      '�Ծ����                       
application("kssj")=rs1("kssj")   '����ʱ��                         
application("kshm")=rs1("kshm")  '���Ժ���                     
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
<center>��ʦ��û���Ծ������ĵȴ�.....
<br><br><p> <a href=kh.asp>-----����-----</a></p></center>
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
      <p align="center"><font color="#FFFFFF">ѧ��������Ϣ</font></td>
  </tr>
  <tr>
    <td width="43%" align="right" height="23">������<font color="#FF0000">����</font>:</td>
    <td width="57%" height="23"><input type ="text" name="xm" size="20"></td>
  </tr>
  <tr>
    <td width="43%" align="right" height="23">������<font color="#FF0000">ѧ��</font>:</td>
    <td width="57%" height="23"><input type ="text" name="xh" size="20"></td>
  </tr>
  <tr>
    <td width="43%" align="right" height="15">������<font color="#FF0000">�༶</font>:</td>
    <td width="57%" height="15"><input type ="text" name="bj" size="20"></td>
  </tr>
  <tr>
    <td width="43%" height="6"></td>
    <td width="57%" height="6"><input type="submit" name="submit" value="��ʼ����">
</td>
  </tr>
</table>
</form>
  </center>
</div>
</body>
</html>
<%end if%>