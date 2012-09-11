<!--#include file="conn.asp" -->

<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>untitled document</title>
</head>

<body bgcolor="#ffffff">
 
<% 
dim sm,dh,exec,na,da,key,tx,ts
function zh(fString)

    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, " ", "&nbsp;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "<BR>")
    zh = fString
end function
sm=session("sm")
dh=session("dh")  
key=session("key")  
exec=session("exec")

set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
%> 

 <%tx=""
 ts=0
%>
  <%do while not rs.eof            
ts=ts+1
if tx<>rs("section") then
 ts=1
 tx=rs("section")
 end if
 %>
 <%na=tx&ts                                                               
   da=request.form(na)                                                         
    if zh(da)<> rs("answer") then 
    rs("answer")=zh(da)                                                     
    rs.update 
    end if
    %>                                                            
<%rs.movenext%><%loop%>            

<% 
rs.close
conn.close 
%>

<br><br><center>已成功地保存了答案....
<br><br><center><a href=ask.asp?dh=<%=dh%>&sm=<%=sm%>&key=<%=key%>
>返回到上一页</a>
<p><a href=kh.asp>回到考试首页</a></p></center>
</body>
</html>