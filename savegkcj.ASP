<!--#INCLUDE FILE="conn.asp" -->
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

dim TT1,TT2,TT3,TT4,TT5,TT6,TT7,TT8,TT9

function zh(fString)

    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, " ", "&nbsp;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10), "<BR>")
    zh = fString
end function
TT1=request.form("T1")
TT2=request.form("T2") 
TT3=request.form("T3")
TT4=request.form("T4")
TT5=request.form("T5")
TT6=request.form("T6")
TT7=request.form("T7")
TT8=request.form("T8") 
TT9=request.form("T9")
if TT6="" then TT6=0 else TT6=cINT(TT6)
if TT7="" then TT7=0 else TT7=cINT(TT7)
if TT8="" then TT8=0 else TT8=cINT(TT8)
if TT9="" then TT9=0 else TT9=cINT(TT9)

%> 
<% 
dim exec 
set rs=server.createobject("adodb.recordset")
rs.open "select * from gkcj",conn,3,3
rs.addnew
rs("准考证号")=TT1
rs("姓名")=TT2
rs("报名号")=TT3
rs("班级")=TT4
rs("专业")=TT5
rs("语文")=TT6
rs("数学")=TT7
rs("英语")=TT8
rs("综合")=TT9
rs("总分")=TT6+TT7+TT8+TT9
rs.update 
%> 
<p align="center">已经添加!</p> 
<%response.redirect "addgkcj.asp"%>
<p align="center"><a href="ask1.asp">继续添加</a></p> 
</body> 
 
</html>