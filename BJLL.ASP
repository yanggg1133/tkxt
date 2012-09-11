<!--#INCLUDE FILE="conn.asp" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 4</title>
</head>

<body>
<% dim exec,bj
exec="select bj from xscj order by  bj desc"
'set conn=server.createobject("adodb.connection") 
'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3

%>
<p align="center"><font size="4" face="隶书" color="#000000"><span style="background-color: #FF0066">选择考试科目菜单</span></font></p>
<table border="1" width="100%" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#66FF66" >


<% bj=""

do while not rs.eof%>              
<% if bj<>rs("bj") then
bj=rs("bj")%>
<tr><td>
<a href=cxxt.asp?bj=<%=rs("bj")%>><%=rs("bj")%></a>              
</td></tr>
<%end if%>
<%rs.movenext%><%loop%>             
<%rs.close
  conn.close
%>
</table>
<p><center><a href=cxxt.asp>直接返回</a></center></p>
</body>

</html>
