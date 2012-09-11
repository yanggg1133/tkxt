<!--#INCLUDE FILE="conn.asp" -->
<%response.write "jjjjjjjjjjjj"
response.end
dim que,dif,dh,sm,key,cz,exec
 dh=request.form("dh")
 sm=request.form("sm")
 key=request.form("key")
 que=request.form("section")
 dif=request.form("diff")
 cz=request.form("cz")
 session("section")=que
 session("diff")=dif
session("dh")=dh
 session("sm")=sm
session("key")=key
 session("cz")=cz
 'set conn=server.createobject("adodb.connection") 
'conn.open "dsn=xiong1" 
exec="select * from sjdh where (dh='"&dh&"')"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
 if not rs.eof then
 select case cz
  case "zj"
  response.redirect "ask1.asp"
  case "xg"
  exec="upda sjdh  where  dh='"+&dh+"'"+" set sm='"+&sm+"',key='"+&key+"'"
  case "cs"
  response.redirect "ask.asp"
  case "sc1"
  
  case "sc2"
  
end select  

 %>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 4</title>
</head>

<body>

</body>

</html>
<%rs.close%>
<%conn.close%>