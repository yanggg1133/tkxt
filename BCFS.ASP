<!--#INCLUDE FILE="conn.asp" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>

<body>
<%dim xh,kshm,xc,tk,pd,js,jd,cx,wxtk,cxjg
xh=session("xh")
kshm=session("kshm")
xc=cint(request.form("选择题"))
tk=cint(request.form("填空题"))
pd=cint(request.form("判断题"))
js=cint(request.form("计算题"))
jd=cint(request.form("简答题"))
cx=cint(request.form("程序题"))
wxtk=cint(request.form("完型填空"))
cxjg=cint(request.form("程序结果"))
if tk=""  then tk=0                     
if xc=""  then xc=0                     
if pd=""  then pd=0                     
if js=""  then js=0                     
if jd=""  then jd=0                     
if cx=""  then cx=0                     
if wxtk="" then wxtk=0                     
if cxjg="" then cxjg=0                     

'set conn=server.createobject("adodb.connection")                      
'conn.open "dsn=xiong1"                      
set rs=server.createobject("adodb.recordset")                     
rs.open "select * from xscj where(kshm='"&kshm&"' and xh='"&xh&"')" ,conn,3,3                     

if not rs.eof then
rs("xc")=xc                     
rs("tk")=tk                     
rs("pd")=pd                     
rs("js")=js                     
rs("jd")=jd                     
rs("cx")=cx                     
rs("wxtk")=wxtk                     
rs("cxjg")=cxjg                     
rs("zf")=xc+tk+pd+js+jd+wxtk+cxjg+cx
rs("gf")=1
rs.update                     
end if
rs.close
conn.close

%>                     
<center><br><br><p> <a href=modify2.asp>成功追加,评下一个的试卷</a> </p>
<br><p> <a href=modify.asp>选择另一套试卷</a> &nbsp;&nbsp;&nbsp;
<br> <a href=index.htm>不改卷了，回到首页</a>
    </center>                     

</body>

</html>
