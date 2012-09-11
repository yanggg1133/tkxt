<!--#INCLUDE FILE="conn.asp" -->
<%
dim id,pwd,ask
ask=request.querystring("ask")
id=request.form("id")
pwd=request.form("pwd")
if id="" then id=session("id")
if  pwd="" then pwd=session("pwd")
session("id")=id
session("pwd")=pwd
%>

<%
 dim exec
exec="select * from tblusr where (usrid='"&id&"'and usrpwd='"&pwd&"')"
'set conn=server.createobject("adodb.connection")
 'conn.open "dsn=xiong1"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
'set rs=conn.execute( "select * from tblusr where (usrid='"&id&"'and usrpwd='"&pwd&"')" )
 
 %>

<%
dim dh,sm,key
dh=request("dh")
sm=request("sm")
key=request("key")
if not rs.eof then
%>
<% if ask="xtgl" then
session("userna")="xtgl"
response.redirect "xtgl.asp"
 end if%>
<script language="vbscript">
sub button1_onclick
 msgbox "可以通过浏览来选择已有的试卷代号和说明以及关键词"
 end sub
 </script> 

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
<script language="javascript">
function check(myform)
{
if(myform.question.value=="")
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
function disp(x)
{
if(x==1)
document.all.xs.style.display="block"
else
document.all.xs.style.display="none"
}
</script>

 <title>验证错误!</title>

</head>

<body bgcolor="#ffffff">
<form method="post" action="ask0.asp"  name="myform" >
<center>
<p>输入试卷代号：<input type="text" name="sjdh" size="10" maxlength="10" value='<%=dh%>'> </p>
<p> <input type="button" name="button1" value="提示" >----><a href=ll.asp?asp=ask>浏览 </a> </p>
  <p>输入说明：<input type="text" name="sjsm" size="80" maxlength="80" value='<%=sm%>'> </p>
  <p>关键词是：<input type="text" name="sjkey" size="80" maxlength="80" value='<%=key%>'> </p>
<p>相同试卷代号处理：
<!--<input type="radio" name="cl" value="zj"> -->
<input type="radio" onclick="disp(1)" name="cz" value="zj" checked> 追加试题&nbsp;&nbsp;
<input type="radio" onclick="disp(0)" name="cz" value="cs" >重新输入代号&nbsp; 
<input type="radio" onclick="disp(0)" name="cz" value="xg">改说明关键词&nbsp;
<input type="radio" onclick="disp(1)" name="cz" value="sc1">只删该类型题</p>
&nbsp;&nbsp;
<input type="radio" onclick="disp(0)" name="cz" value="sc2">删除该试卷&nbsp;
<input type="radio" onclick="disp(0)"  name="cz" value="da">输入（修改）答案&nbsp;&nbsp;

<input type="radio" onclick="disp(1)" name="cz" value="xgst">修改试题&nbsp;&nbsp; <div id=xs style="display:block"> 
<p><span style="background-color: #C0C0C0"><font color="#0000FF">请选择题型:      
<select name="section">      
<option>选择题</option> 
<option>填空题</option> 
<option>判断题</option> 
<option>计算题</option> 
<option>简答题</option> 
<option>程序题</option> 
<option>完型填空</option> 
<option>程序结果</option> 
</select> </font></span> </p>
<p><span style="background-color: #C0C0C0"><font color="#0000FF">请选择难度:      
<input type="radio" name="diff" value="hard" >      
较难      
<input type="radio" name="diff" value="normal" checked>      
中等      
<input type="radio" name="diff" value="easy" >      
容易</font></span></p>   
  </div>
<P>      
<input type="submit" name="submit" value="填好了">      
      
<input type="reset" name="submit2" value="重填一次">      
</p>  
</form> 
</center>  
 <%else%> 
<br><br><br> 
<center>  
<p><a href=login.htm>返回到上一页，输入用户名和密码</a></p>  
<p><a href=index.htm>回到首页</a></p>  
<p align="center">密码错误,请重新登录</p>     
<p align="center"><a href="javascript:history.back()">返回</a></p></center> 
<%     
end if     
%>     
<%rs.close%>     
<%conn.close%> 
</body>  
</html>