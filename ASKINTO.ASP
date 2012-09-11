<!--#INCLUDE FILE="conn.asp" -->
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

dim section,diff,question,choicea,choiceb,choicec,choiced,answer,dh 

'section=request.form("section") 
'section=session("section")
function zh(fString)

    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, " ", "&nbsp;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "<BR>")
    zh = fString
end function
section=session("section")
dh=session("dh")
diff=session("diff")
question=zh(request.form("question"))
choicea=zh(request.form("choicea")) 
choiceb=zh(request.form("choiceb"))
choicec=zh(request.form("choicec"))
choiced=zh(request.form("choiced"))
answer=zh(request.form("answer"))
%> 
<% 
dim exec 
'set conn=server.createobject("adodb.connection") 
'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
rs.open "select * from question where(dh='"+dh+"')",conn,3,3
rs.addnew
rs("dh")=dh
rs("section")=section
rs("diff")=diff
rs("que")=question
rs("choicea")=choicea
rs("choiceb")=choiceb
rs("choicec")=choicec
rs("choiced")=choiced
rs("answer")=answer
rs.update 
%> 
<p align="center">您的题目已经添加!</p> 
<p align="center"><a href="ask1.asp">继续添加</a></p> 
<br><br><br>
<center>
<p>  
<a  href=ask.asp?dh=<%=session("dh")%>&sm=<%=session("sm")%>&key=<%=session("key")%>>返回到上一页，继续选择题型和难度</a>  
</p>   
 <p><a href=index.htm>回到首页</a></p>
 </center>
</body> 
 
</html> 
<%rs.close
conn.close%> 
