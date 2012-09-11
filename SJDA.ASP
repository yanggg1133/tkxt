<!--#include file="conn.asp" -->

<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>untitled document</title>
</head>

<body bgcolor="#ffffff">
 
<% 
dim i,j,kssj,kshm,sm,dh,exec,m,xh,xm,bj,tx(8,2),na,da
function zh(fString)

    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, " ", "&nbsp;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "<BR>")
    zh = fString
end function
xh=request.form("xh")
xm=request.form("xm")
bj=request.form("bj")
kssj=application("kssj")
kshm=application("kshm")
dh=application("dh")
sm=application("sm")
exec=application("exec")
m="select * from xscj where(kshm='"+kshm+"' and xh='"+xh+"')"
'set conn=server.createobject("adodb.connection") 
'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
rs.open m,conn,3,3
if  rs.recordcount=0 then
rs.addnew
rs("dh")=dh           
rs("kshm")=kshm           
rs("xm")=xm      
rs("xh")=xh           
rs("bj")=bj           
rs("exec")=exec
rs.update              
else
rs.close
conn.close
response.redirect "kh.asp"
end if
%> 

 <%m="select * from sj where(kshm='"+kshm+"' and dh='"+dh+"')"
 rs.close
 rs.open m,conn,3,3
 if rs.recordcount=0 then %>
 <%rs.close
 conn.close%>
 <center>
<p>本套试卷不能被保存...请与发卷人联系....</p>
 <p><a href=kh.asp>回到考试首页</a></p></center>
 <%else
  tx(0,1)=rs("xc")
  tx(1,1)=rs("wxtk")
  tx(2,1)=rs("tk")
  tx(3,1)=rs("pd")
  tx(4,1)=rs("jd")
  tx(5,1)=rs("js")
  tx(6,1)=rs("cx")
  tx(7,1)=rs("cxjg")
rs.close
tx(0,0)="选择题"
tx(1,0)="完型填空"
tx(2,0)="填空题"
tx(3,0)="判断题"
tx(4,0)="简答题"
tx(5,0)="计算题"
tx(6,0)="程序题"
tx(7,0)="程序结果"
m="select * from da where(kshm='"+kshm+"' and xh='"+xh+"')"
rs.open m,conn,3,3
for i=0 to 7
  for j=1 to tx(i,1)
   na=tx(i,0)&j
   da=request.form(na)
  rs.addnew
  rs("kshm")=kshm
  rs("dh")=dh
  rs("xh")=xh
  rs("da")=zh(da)
  rs("tx")=na
  rs.update
next
next
rs.close
conn.close 
%>

<br><br><center>已成功地保存了答案....
<p><a href=kh.asp>回到考试首页</a></p></center>

<%end if
%>
</body>
</html>