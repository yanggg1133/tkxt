<!--#INCLUDE FILE="conn.asp" -->
<%
dim id,pwd,teacher
teacher=request.querystring("teacher")
id=request.form("id")
pwd=request.form("pwd")
%>

<%
dim exec
exec="select * from tblusr where (usrid='"&id&"'and usrpwd='"&pwd&"')"
'set conn=server.createobject("adodb.connection")
'conn.open "dsn=xiong1"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
%>

<%
if not rs.eof then
 if teacher="input" then
response.redirect "teacher1.asp"
 else
 response.redirect "modify.asp"
 end if
%>
<%else%>


<html>
<head> 
   
 <title>验证错误!</title>
</head>
<body>


<p align="center">密码错误,请重新登录</p>    
<p align="center"><a href="javascript:history.back()">返回</a></p>   
</body> 
</HTML> 
 
<%    
end if    
%>    
<%rs.close%>    
<%conn.close%>









