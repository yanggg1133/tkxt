<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>
<body>

<%dim hm,value,pp,exec,hm1
  dim zz(5)
  zz(1)="姓名"
  zz(2)="准考证号"
  zz(3)="班级"
  zz(4)="报名号"
  zz(5)="专业"
hm=request.form("item")
hm1=zz(cINT(hm))
value=trim(request.form("va"))
set rs=server.createobject("adodb.recordset")  
%>                  
 

 
<%
if value<>"" then 
  if value="*" then
  exec= "delete  from gkcj"
  else
  exec= "delete  from gkcj where trim("&hm1&")='"&value&"'"  
  end if
  conn.execute exec
  %>
  <p align="center">成功删除</p>
<br>  
<a align="center" href="delegkcj.asp">返回</a>
<%else%>
<p align="center">未删除任何记录</p>
<br>  <a align="center" href="delegkcj.asp">返回</a>
   <%
   response.end 
  end if 
%>

  
</body> 
 </html>