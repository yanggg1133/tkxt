
<!--#INCLUDE FILE="conn.asp" -->

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
</head>

<body>

<%
dim id,pwd
id=request.form("user")
pwd=request.form("pass")
%>

<%
 dim exec
exec="select * from tblusr where (usrid='"&id&"'and usrpwd='"&pwd&"')"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
 
if  rs.eof then
%>
<br>密码错，返回....<br>
<a href="userlogin.asp">返回</a>
<%
else
%>

<p><b><span style="background-color: #FFFF00">成绩管理操作</span></b></p>
<p><span style="background-color: #00CCFF"><b><font size="5" color="#0000FF">
<a href="addgkcj.asp" target="right">增加记录</a></font></b></span></p>
<p><span style="background-color: #00CCFF"><font size="5" color="#0000FF"><b>
<a href="modigkcj.asp?jlh=0" target="right">修改记录</a></b></font></span></p>
<p><span style="background-color: #00CCFF"><font size="5" color="#0000FF"><b>
<a href="delegkcj.asp" target="right">删除记录</a></b></font></span></p>
<p><span style="background-color: #808080"><font size="5" color="#FFFF66"><b>
<a href="userlogin.asp">返回查询</a></b></font></span></p>
<%end if%>
</body>

</html>