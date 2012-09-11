<!--#INCLUDE FILE="conn.asp" -->


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>asdfasdfasdfasdfasdfasdfasdf</title>
</head>

<body>
<% 

dim exec 
asp=request("asp")
exec="select * from sjdh"
'set conn=server.createobject("adodb.connection") 
'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
%>
<p><center>本题库一共有：<%=rs.recordcount%>&nbsp;&nbsp;&nbsp;套题可以选择</center></p>
<%if asp="ask" then%>
<a href=ask.asp><p><center>试卷代号 &nbsp;&nbsp;&nbsp;     说明 &nbsp;&nbsp;&nbsp;       关键词</center></p><br></a>                  
<%else%>           
<a href=teacher1.asp><p><center>试卷代号 &nbsp;&nbsp;&nbsp;     说明 &nbsp;&nbsp;&nbsp;       关键词</center></p><br></a>                  
<%end if%>           
<% if not rs.eof then%>                  
<p align="center">                  
<%do while not rs.eof%>                  
<%if asp="ask" then%>           
<a href=ask.asp?dh=<%=rs("dh")%>&sm=<%=rs("sm")%>&key=<%=rs("key")%>><%=rs("dh")%>;&nbsp;&nbsp;<%=rs("sm")%>;&nbsp;&nbsp;<%=rs("key")%><br></a>                  
<%else%>           
<%'session("dh")=rs("dh")         
'session("sm")=rs("sm")         
'session("key")=rs("key")%>         
<a href=teacher1.asp?dh=<%=rs("dh")%>&sm=<%=rs("sm")%>&key=<%=rs("key")%>><%=rs("dh")%>;&nbsp;&nbsp;<%=rs("sm")%>;&nbsp;&nbsp;<%=rs("key")%><br></a>                  
<%end if%>             
<%rs.movenext%><%loop%></p>                  
                  
<%end if%>                  
<%rs.close%>                  
<%conn.close%>                  
</body>                  
                  
</html>                  
