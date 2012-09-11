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
dim sf,na1,na,sum,dq1,gs,m
 dq1=session("dq1")
 gs=session("gs")

set rs=server.createobject("adodb.recordset")
rs.open "select * from gkcj",conn,3,3
rs.move dq1

 row=1
m=0
 do while not rs.eof  and row<=gs
 
 na1="tt"&row&"10"
 sf=request.form(na1)                                                         
 if sf="ON"  then
 m=m+1
 k=1
 sum=0
 do while k<=9
 na="tt"&row&k
 TT1=request.form(na)
if k=6 or k=7 or k=8 or k=9 then
if TT1="" then TT1=0 else TT1=cINT(TT1)
sum=sum+TT1
end if
 rs(k-1)=tt1
 k=k+1
 loop
rs(9)=sum 
rs.update 
end if
row=row+1                                                       
rs.movenext%><%loop %>            

<br><br><center>已成功地保存了记录...<b><center><span style="background-color: #0000FF">
<font size="6">共保存了:<%=m%>条记录</font></span></center>
<br><br><center><a href=modigkcj.asp?jlh=-1
>继续修改</a>
<p><a href=kh.asp>回到考试首页</a></p></center>
</body>
</html>