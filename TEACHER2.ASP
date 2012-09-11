<!--#INCLUDE FILE="conn.asp" -->
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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>asdfasdfasdfasdfasdfasdfasdf</title>
</head>

<body>
<% 

dim exec ,sf,dh,tx,diff,diff1,diff2,diff3,ts,sf1,sm,key,tk,xc,pd,js,jd,cx,wxtk,cxjg,m,kshm
kshm=request.form("kshm")
sm=session("sm")
dh=request.form("sjdh")
diff1=request.form("diff1")
diff2=request.form("diff2")
diff3=request.form("diff3")
sf=request.form("sf")
sf1=request.form("sf1")
if sf1="yesno" then 
application("exec")=""
response.redirect "teacher1.asp"
end if
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
  exec="select * from question where (("+tx+") and "+diff+" and dh='"+dh+"')"+"order by section desc"
  
  session("exec")=exec

  %>
  <%if sf1="no" then %>
 
   <%

'set conn=server.createobject("adodb.connection") 

'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
%>
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
    <%=tx%>　</td>
 </tr>
 <%end if%>
 <tr width="100%" ><td colspan="2"><%=ts&"."%><%=rs("que")%>　</td>
 </tr>
 <%if tx="选择题" then%>
 <tr align="center"><td width="50%">A:<%=rs("choicea")%></td><td width="%50">B:<%=rs("choiceb")%></td></tr>
  <tr align="center"><td width="50%">C:<%=rs("choicec")%></td><td width="50%">D:<%=rs("choiced")%></td></tr>
  <%end if%>
<%rs.movenext%><%loop%>            
</table>
<%rs.close
  %>
<br><br><center><a href=teacher1.asp>返回到上一页</a> </center>
<%else%>


<%                                 
'set conn=server.createobject("adodb.connection")                                  
'conn.open "dsn=xiong1"                                  
set rs=server.createobject("adodb.recordset")                                 

rs.open "select * from sj where(status=1)" ,conn,1,3                                 
if not rs.eof then 
 rs("status")=0
rs.update
rs.close
end if

set rs=server.createobject("adodb.recordset")
rs.open "select * from sj where(kshm='"&kshm&"')" ,conn,1,3                                 

   if not rs.eof then                 
   application("exec")=rs("tx")   '考试试题                          
application("dh")=rs("dh")      '试卷代号                       
application("sm")=rs("sm")      '试卷标题                       
application("kssj")=rs("kssj")   '考试时间                         
application("kshm")=rs("kshm")  '考试号码                     
  

rs("status")=1
rs.update
 
   
   %>               
   <center><br><br><p> <a href=teacher1.asp?kshm=<%=rs("kshm")%>>选择了已存在的一套试卷,返回到上一页</a> </p>               
                  
<p>标题是:<%=rs("sm")%></p>    </center>                                 
               
   <%else                         
tk=session("tk")                                 
xc=session("xc")                                 
pd=session("pd")                                 
js=session("js")                                 
jd=session("jd")                                 
cx=session("cx")                                 
wxtk=session("wxtk")                                 
cxjg=session("cxjg")                                 
dh=session("dh")                                 
sm=session("sm")                                 
if tk="" or request("tk")="" then tk=0                                 
if xc="" or request("xc")="" then xc=0                                 
if pd="" or request("pd")="" then pd=0                                 
if js="" or request("js")="" then js=0                                 
if jd=""  or request("jd")="" then jd=0                                 
if cx="" or request("cx")="" then cx=0                                 
if wxtk=""  or request("wxtk")="" then wxtk=0                                 
if cxjg=""  or request("cxjg")="" then cxjg=0                                 
'if tk="" or xc="" or pd="" or js="" or jd="" or cx="" or wxtk="" or cxjg="" then   
'msgbox "未选择浏览，不能被确定为考试卷！"   
'response.redirect "teacher1.asp"   
'endif           
rs.addnew                                 
rs("dh")=dh                                 
rs("sm")=sm                                 
rs("kssj")=request.form("kssj")                            
rs("xc")=xc                                 
rs("tk")=tk                                 
rs("pd")=pd                                 
rs("js")=js                                 
rs("jd")=jd                                 
rs("cx")=cx                                 
rs("wxtk")=wxtk                                 
rs("tx")=exec                                 
rs("kshm")=request.form("kshm")                             
rs("cxjg")=cxjg                                 
rs("status")=1

rs.update                                 
%>                                 
<center><br><br><p> <a href=teacher1.asp>成功追加,返回到上一页</a>     </center>                                 
<%application("exec")=exec   '考试试题                          
application("dh")=dh      '试卷代号                       
application("sm")=sm      '试卷标题                       
application("kssj")=request.form("kssj")   '考试时间                         
application("kshm")=request.form("kshm")  '考试号码                        
end if%>                                 
                 
<%end if%>                                   
                                   
<center><br>　 <a href=index.htm>回到首页</a>                                     
</center>                                   
</body>                                                  
                                                  
</html>