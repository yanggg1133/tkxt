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
{alert("请输入账号");
 return(false);
 }
 if(myform.pwd.value=="")
 {alert("请输入密码");
 return(false);
 }}
 </script>
 
</head>

<body align="center" bgcolor="#ffffff">

<%
dim modify
modify=request.querystring("teacher")
%>
<p>　</p>
<form name="form1" method="post" action="teacher.asp?teacher=<%=modify%>" onsubmit="return check(this)">
<center>
<p>&nbsp;  
 </p>  
<table border="1" width="37%" height="135" cellpadding="2">
  <tr>
    <td width="100%" colspan="2" bgcolor="#0000FF" height="16">
    <p align="center"><font color="#FFFFFF"><%if modify="input" then%>出题教师登录<% else%>改卷教师登录<%end if%></font></p></td>
  </tr>
  <tr>
    <td width="42%" height="23">
      <p align="right">请输入<font color="#FF0000">账号</font>:</p>
    </td>
    <td width="54%" height="23">  
 <input type ="text" name="id">  
    </td>
  </tr>
  <tr>
    <td width="42%" height="25">
      <p align="right">请输入<font color="#FF0000">密码</font>:</p>
    </td>
    <td width="54%" height="25">  
  <input type="password" name="pwd">  
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="1">
      <p align="center">
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="submit" name="submit" value="提交">    
      </p>  
    </td>  
  </tr>  
  <tr>  
    <td width="100%" colspan="2" height="1">  
      <p align="center"><font color="#FF00FF">　<a href="index.htm"><font size="4" face="华文琥珀" color="#FF00FF">首  
      页</font></a><font size="4" face="华文琥珀">&nbsp;&nbsp;&nbsp; </font></font><a href="kh.asp"><font size="4" face="华文琥珀" color="#FF00FF">上一页</font></a>  
      </p>  
    </td>  
  </tr>  
</table>  
 <p>&nbsp;    
  </p>    
  <p>    
  　  
  </p>  
</center>  
  </form>  
  <p>&nbsp;</p> 
<p align="center">  
 &nbsp;&nbsp;&nbsp;&nbsp;</p>  
  
</body>  
  
</html>  
