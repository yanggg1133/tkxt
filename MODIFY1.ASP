<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>考试号码的浏览</title>
</head>

<body>

<p align="center"><font size="5" face="隶书">选择待评试卷的考试号码</font></p>
<p align="center"><font size="3" face="方正舒体">你必须先选择已经存在的考试号码才能评卷</font></p>

<% dim kshm
kshm=request.querystring("kshm")                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")                    
rs.open "select * from sj where(kshm='"&kshm&"')",conn,3,3                    
  if not rs.eof then     
   session("exec")=rs("tx")   '考试试题              
session("dh")=rs("dh")      '试卷代号           
session("sm")=rs("sm")      '试卷标题           
session("kssj")=rs("kssj")   '考试时间             
session("kshm")=rs("kshm")  '考试号码         
response.redirect "modify2.asp"  
 else%> 
<center>不存在该试卷,请返回....</center><p>
<a href="kh.asp"><span style="background-color: #FF6666"><font size="4" face="华文琥珀" color="#66FF33">返回到上一页</font></span></a>
 </p>

<%end if
rs.close
conn.close%>
</body>

</html>









