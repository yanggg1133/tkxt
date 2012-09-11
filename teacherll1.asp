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

dim g,exec ,dh,tx(8),diff,diff1,diff2,diff3,xc,tk,pd,jd,js,cx,wxtk,cxjg
xc=""
tk=""
js=""
pd=""
jd=""
cx=""
wxtk=""
cxjg=""
dh=request("dh")
diff1=request("diff1")
diff2=request("diff2")
diff3=request("diff3")
diff=""
diff="(diff='"+diff1+"' or diff='"+diff2+"' or diff='"+diff3+"')"
tx(0)="选择题"
tx(1)="填空题"
tx(2)="判断题"
tx(3)="计算题"
tx(4)="简答题"
tx(5)="程序题"
tx(6)="完型填空"
tx(7)="程序结果"

'set conn=server.createobject("adodb.connection") 

'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")

for i=0 to 7
exec="select dh from question where (section='"+tx(i)+"' and "+diff+"and dh='"+dh+"')"
rs.open exec,conn,3,3
g= rs.recordcount
if g>0 then
select case i
case 0
 xc=g
case 1
 tk=g
case 2
 pd=g
 case 3
 js=g
 case 4
 jd=g
 case 5
 cx=g
 case 6
 wxtk=g
 case 7
 cxjg=g
 end select

end if
rs.close
'response.write exec1
next
session("dh")=dh
session("diff1")=diff1
session("diff2")=diff2
session("diff3")=diff3

session("xc")=xc
session("tk")=tk
session("pd")=pd
session("js")=js
session("jd")=jd
session("cx")=cx
session("wxtk")=wxtk
session("cxjg")=cxjg

response.redirect "teacher1.asp"

%>
<%%>
<%conn.close%>              
</body>              
              
</html>              
