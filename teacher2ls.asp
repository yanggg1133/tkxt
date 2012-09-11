<%@language=vbscript %>
<script language="javascript">
function check(myform)
{if(myform.question.value=="")
{alert("请输入题目");
return false;
}}
</script>
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

dim exec ,sf,dh,tx,diff,diff1,diff2,diff3,ts,sf1,na,m
dh=request.form("sjdh")
diff1=request.form("diff1")
diff2=request.form("diff2")
diff3=request.form("diff3")
sf=request.form("sf")
sf1=request.form("sf1")
diff="(diff='"+diff1+"' or diff='"+diff2+"' or diff='"+diff3+"')"
if sf="yes"  then
response.redirect "teacherll1.asp?dh="&dh&"&diff1="&diff1&"&diff2="&diff2&"&diff3="&diff3
end if
  tx="section=''"
  if request.form("xc")<>"" then
   tx=tx+" or section='"+request.form("xc")+"'"
   end if
   if request.form("tk")<>"" then
   tx=tx+" or section='"+request.form("tk")+"'"
   end if
    if request.form("pd")<>"" then
   tx=tx+" or section='"+request.form("pd")+"'"
   end if
    if request.form("js")<>"" then
   tx=tx+" or section='"+request.form("js")+"'"
   end if
    if request.form("jd")<>"" then 
   tx=tx+" or section='"+request.form("jd")+"'"
   end if
   
    if request.form("cx")<>"" then
   tx=tx+" or section='"+request.form("cx")+"'"
   end if
    if request.form("wxtk")<>"" then
   tx=tx+" or section='"+request.form("wxtk")+"'"
   end if
    if request.form("cxjg")<>"" then
   tx=tx+" or section='"+request.form("cxjg")+"'"
   end if
     
    %>
   <%

set conn=server.createobject("adodb.connection") 

conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
exec="select * from question where (("+tx+") and "+diff+" and dh='"+dh+"')"+"order by section desc"
session("exec")=exec
rs.open exec,conn,3,3
%>
<%if sf1="yes" then %>
<form method="post" action="sjda.asp"  name="myform" onsubmit ="return check(this)">
<%end if%>
<center><span style="background-color: #FFFF00"><font face="楷体_GB2312" color="#0000FF" size="5"><%=session("sm")
%></font></span>


<center>
<table align="lift" border="1" width="100%">
<%tx=""
 ts=0
%>
  <%do while not rs.eof            
ts=ts+1
if tx<>rs("section") then
 ts=1
 tx=rs("section")
 %>
 <tr align="center"><td  width="100%" colspan="2">
    <%=tx%></td>
 </tr>
 <%end if%>
 <tr width="100%" ><td colspan="2"><%=ts&"."%><%=rs("que")%></td>
 </tr>
 <%if tx="选择题" then%>
 <tr align="center"><td width="50%">A:<%=rs("choicea")%></td><td width="%50">B:<%=rs("choiceb")%></td></tr>
  <tr align="center"><td width="50%">C:<%=rs("choicec")%></td><td width="50%">D:<%=rs("choiced")%></td></tr>
  <%end if%>
  
  
 <%if sf1="yes" then
  na=tx&ts %>
  <tr><td width="100%"  colspan="2">
<% if tx="选择题"then
  m="请选择：<input type='radio' name='"+na+"' value='A'> A"%>
  <%=m%>
 <% m="<input type='radio' name='"+na+"' value='B'> B"%>
<%  =m%>
  <%m="<input type='radio' name='"+na+"' value='C' checked> C"%>
 <% =m     %>
<%m="<input type='radio' name='"+na+"' value='D'> D"%>
 <% =m %>
  <%else 
  if tx="判断题" then
  m="正确否：<input type='radio' name='"+na+"' value='yes'  checked>正确"%>
  <%=m%>
 <% m="<input type='radio' name='"+na+"' value='no'>错误"%>
<%  =m%>
  
<%else 
  m="答案是：<textarea  name='"+na+"' cols='95' rows='2'></textarea>"%>
  <%=m%>
  <%end if%>
  <%end if%></td></tr>
  <%end if%>
<%rs.movenext%><%loop%>            
</table>
  <p><input type="submit" name="submit" value="交卷">     </p>
<% if sf1="yes" then %>
 </form>
 <%end if%>　
<p> <a href=choose.asp>放弃并返回</a>     

　 <a href=index.htm>回到首页</a>     </p>
</body>               
               
</html>