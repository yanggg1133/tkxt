<!--#INCLUDE FILE="conn.asp" -->
<%
dim id,pwd
id=""
pwd=""
id=request.form("id")
pwd=request.form("pwd")
%>

<%
 dim exec
exec="select * from tblusr where (usrid='"&id&"'and usrpwd='"&pwd&"')"
exec="select * from tblusr"
 'set conn=server.createobject("adodb.connection")
 'conn.open  "driver={microsoft  access driver (*.mdb)};dbq="&dbpath
 'conn.open "dsn=xiong1"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
'set rs=conn.execute( "select * from tblusr where (usrid='"&id&"'and usrpwd='"&pwd&"')" )
 if not rs.eof then
 %>
 
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>输入题目</title>
<script language="javascript">
function check(myform)
{if(myform.question.value=="")
{alert("请输入题目");
return false;
}
if(myform.choicea.value=="")
{alert("请输入选项A");
return false;
}
if(myform.choiceb.value=="")
{alert("请输入选项B");
return false;
}
if(myform.choicec.value=="")
{alert("请输入选项c");
return false;
}
if(myform.choiced.value=="")
{alert("请输入选项d");
return false;
}
}
</script>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">


</head>

<body bgcolor="#ffffff">
<form method="post" action="askinto.asp"  name="myform" onsubmit ="return check(this)">
<p>请选择题型:    
<select name="section">    
<option>选择题</option>
<option>填空题</option>
<option>判断题</option>
<option>计算题</option>
<option>简答题</option>
<option>程序题</option>
<option>完型填空</option>
<option>程序结果</option>
</select>
</p>
<p>请选择难度:    
<input type="radio" name="diff" value="hard" >    
较难    
<input type="radio" name="diff" value="normal" checked>    
中等    
<input type="radio" name="diff" value="easy" >    
容易</p>    
<p>请输入问题:    
<input type="text" name="question" size="50" maxlength="50">    
(25字以内)</p>    
<p>请输入A选项:    
<input type="text" name="choicea" size="50" maxlength="50">    
(25字以内)</p>    
<p>请输入b选项:    
<input type="text" name="choiceb" size="50" maxlength="50">    
(25字以内)</p>    
<p>请输入c选项:    
<input type="text" name="choicec" size="50" maxlength="50">    
(25字以内)</p>    
<p>请输入d选项:    
<input type="text" name="choiced" size="50" maxlength="50">    
(25字以内)</p>    
<p>答案是:    

<input type="radio" name="answer" value="a" checked>    
A    
<input type="radio" name="answer" value="B" >    
B    
<input type="radio" name="answer" value="C" >    
C    
<input type="radio" name="answer" value="D"  >    
D    
</P>
<P>    
<input type="submit" name="submit" value="填好了">    
    
<input type="reset" name="submit2" value="重填一次">    
</p>    
</form> 
</body>
</html>

<%
else
%>
<html>
<head> 
   
 <title>验证错误!</title>
</head>
<body>


<p align="center">密码错误,请重新登录</p>    
<p align="center"><a href=".\login.htm">返回</a></p>   
</body> 
</HTML> 
 
<%    
end if    
%>    
<%rs.close%>    
<%conn.close%>
































































