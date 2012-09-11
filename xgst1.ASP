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
dim exec,dh,sm,key,m,kshm,tx,ts,sf,aa,bb,cc,dd,m1,na1,na,qu,que1,sum
function zh(fString)

    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, " ", "&nbsp;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "<BR>")
    zh = fString
end function
sum=0
qu=session("section")
sm=session("sm")
dh=session("dh")  
key=session("key")  
exec="select * from question where(dh='"+dh+"'and section='"+qu+"')"+"order by section desc"
session("exec")=exec
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3%> 

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
   na1="sf"&ts                                                               
   sf=request.form(na1)                                                         
   if sf="on"  then
     
     que1=request.form(na) 
     if tx="选择题" then
     aa=request.form(na&"a")
     bb=request.form(na&"b")
     cc=request.form(na&"c")
     dd=request.form(na&"d")
     rs("choicea")=aa
     rs("choiceb")=bb
     rs("choicec")=cc
     rs("choiced")=dd
     end if
 
     rs("que")=que1
    rs.update 
   sum=sum+1
   end if
    %>                                                            
<%rs.movenext%><%loop%>            

<% 
rs.close
conn.close 
%>

<br><br><center>已成功地保存了试题...<b><center><span style="background-color: #0000FF">
<font size="6">共保存了:<%=sum%>道题</font></span></center>
<br><br><center><a href=ask.asp?dh=<%=dh%>&sm=<%=sm%>&key=<%=key%>
>返回到上一页</a>
<p><a href=kh.asp>回到考试首页</a></p></center>
</body>
</html>