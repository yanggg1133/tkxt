<!--#include file="conn.asp" -->
<%  
dim que,dif,dh,sm,key,cz 
dh=request.form("sjdh")  
 sm=request.form("sjsm")  
 key=request.form("sjkey")  
 que=request.form("section")  
 dif=request.form("diff")  
 cz=request.form("cz")  
 session("section")=que  
 session("diff")=dif  
session("dh")=dh  
 session("sm")=sm  
session("key")=key  
 'session("cz")=cz  
 %>
 
<% dim exec,m
 if dh="" then
 response.redirect "ask.asp"
 end if
 'set conn=server.createobject("adodb.connection")   
'conn.open "dsn=xiong1"   
exec="select * from sjdh where (dh='"&dh&"')"  
set rs=server.createobject("adodb.recordset")  
rs.open exec,conn,3,3  
 if not rs.eof then  
   select case cz  
  case "cs"  
  m="ask.asp?dh="&dh&"&sm="&sm&"&key="&key
  response.redirect  m     
  rs.close   
  case "zj"     
  response.redirect "ask1.asp"     
  rs.close   
  case "xg"     
 ' exec="update sjdh  set sm='"+sm+"' where  dh='"+dh+"'"   
 'rs.close   
 'set rs=nothing   
 'conn.execute exec   
 'exec="update sjdh  set key='"+key+"' where  dh='"+dh+"'"   
 'conn.execute exec   
' rs.open exec,conn,3,3   
 rs("sm")=sm   
 rs("key")=key   
 rs.update   
rs.close   
response.write "<br><br><br><br><br><center><b>已成功地修改了说明和关键词</b></center>"  
case "sc1"   
exec="delete from question  where dh='"+dh+"' and section='"+que+"'"   
rs.close   
set rs=nothing   
conn.execute exec   
response.write "<br><br><br><br><br><center><b>已成功地删除指定试题</b></center>"   
case "sc2"   
   
rs.close   
set rs=nothing   
   
exec="delete from question  where dh='"+dh+"'"   
conn.execute exec   
exec="delete from sjdh  where dh='"+dh+"'"   
conn.execute exec   
response.write "<br><br><br><br><br><center><b>已成功地删除指定试题</b></center>"   

case "da"
response.redirect "da.asp"
case "xgst"
response.redirect "xgst.asp"
end select      
else   
exec ="insert into sjdh(dh,sm,key) values('"+dh+"','"+sm+"','"+key+"')"    
conn.execute exec   
response.redirect "ask1.asp"     
rs.close   
end if     
 %> 
 <html>
<head>
<meta http-equiv="refresh" content="2;url=ask.asp?dh=<%=session("dh")%>&sm=<%=session("sm")%>&key=<%=session("key")%>">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 4</title>
</head>
<body>
</body>     
 </html>     
<%conn.close%>