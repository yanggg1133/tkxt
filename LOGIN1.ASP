<%@language=vbscript %>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
<script language="javascript">
function check(myform)
{if(myform.id.value=="")
{alert("�������˺�");
 return(false);
 }
 if(myform.pwd.value=="")
 {alert("����������");
 return(false);
 }}
 </script>
 
</head>

<body align="center" bgcolor="#ffffff">

<%
dim modify
modify=request.querystring("teacher")
%>
<p>��</p>
<form name="form1" method="post" action="teacher.asp?teacher=<%=modify%>" onsubmit="return check(this)">
<center>
<p>&nbsp;  
 </p>  
<table border="1" width="37%" height="135" cellpadding="2">
  <tr>
    <td width="100%" colspan="2" bgcolor="#0000FF" height="16">
    <p align="center"><font color="#FFFFFF"><%if modify="input" then%>�����ʦ��¼<% else%>�ľ���ʦ��¼<%end if%></font></p></td>
  </tr>
  <tr>
    <td width="42%" height="23">
      <p align="right">������<font color="#FF0000">�˺�</font>:</p>
    </td>
    <td width="54%" height="23">  
 <input type ="text" name="id">  
    </td>
  </tr>
  <tr>
    <td width="42%" height="25">
      <p align="right">������<font color="#FF0000">����</font>:</p>
    </td>
    <td width="54%" height="25">  
  <input type="password" name="pwd">  
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="1">
      <p align="center">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="submit" name="submit" value="�ύ">    
      </p>  
    </td>  
  </tr>  
  <tr>  
    <td width="100%" colspan="2" height="1">  
      <p align="center"><font color="#FF00FF">��<a href="index.htm"><font size="4" face="��������" color="#FF00FF">��  
      ҳ</font></a><font size="4" face="��������">&nbsp;&nbsp;&nbsp; </font></font><a href="kh.asp"><font size="4" face="��������" color="#FF00FF">��һҳ</font></a>  
      </p>  
    </td>  
  </tr>  
</table>  
 <p>&nbsp;    
  </p>    
  <p>    
  ��  
  </p>  
</center>  
  </form>  
  <p>&nbsp;</p> 
<p align="center">  
 &nbsp;&nbsp;&nbsp;&nbsp;</p>  
  
</body>  
  
</html>  
