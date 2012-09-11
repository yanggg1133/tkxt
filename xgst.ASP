<!--#INCLUDE FILE="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>sdfasdf</title>
<SCRIPT language=JavaScript>
function js(x)
{ 
 xx="document.form1."+x
 eval(xx).checked=1
 
}
function js1(x)
{ 
  
}

</script>
</head>
<SCRIPT language=JavaScript>
var currentpos,timer;
function initialize()
{
timer=setInterval("scrollwindow()",50);
}
function sc(){
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos != document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</script>

<body>
<% 

dim exec,dh,sm,key,m,kshm,tx,ts,sf,sfa,sfb,sfc,sfd,aa,bb,cc,dd,m1,na1,na
function zh(fString)

    fString = replace(fString , "&gt;", ">")
    fString = replace(fString, "&lt;", "<")
    fString = replace(fString, "&nbsp;", " ")
    
    fString = Replace(fString, "<BR>", CHR(13)+chr(10))
    
    zh = fString
end function
qu=session("section")
sm=session("sm")
dh=session("dh")  
key=session("key")  
exec="select * from question where(dh='"+dh+"'and section='"+qu+"')"+"order by section desc"
session("exec")=exec
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
%>
<center><span style="background-color: #FFFF00"><font face="楷体_GB2312" color="#0000FF" size="5"><%=session("sm")
%></font></span>
<form name="form1" method="post" action="xgst1.asp" > 

<table  border="1" width="100%">
<%tx=""
 ts=0
%>
  <%do while not rs.eof            
ts=ts+1
if tx<>rs("section") then
 ts=1
 tx=rs("section")
 %>
 <tr align="center"><td  width="100%" >
    <%=tx%>　</td>
 </tr>
 <%end if%>
 <%na=tx&ts 
   na1="sf"&ts
 %>
 <% m="<textarea onChange='js("+chr(34)+na1+chr(34)+")' name='"+na+"' cols='94' rows='6'>"+zh(rs("que"))+"</textarea>" 
  m1="<input onclick='js1(this)' type='checkbox' name='"+na1+"' value='on'>"
 %>

 <tr  width="100%" ><td  ><%=ts&"."%><%=m%><BR>
   <span lang="en-us">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   </span><span style="background-color: #FF9933">有效否？</span><%=m1%>　</td>
 </tr>
 <%if tx="选择题" then
 
 aa="<input type='text' onChange='js("+chr(34)+na1+chr(34)+")' name='"+tx&ts&"a"+"' value='"+rs("choicea")+"' size='100'>"
  bb="<input type='text' onChange='js("+chr(34)+na1+chr(34)+")' name='"+tx&ts&"b"+"' value='"+rs("choiceb")+"' size='100'>"
 cc="<input type='text' onChange='js("+chr(34)+na1+chr(34)+")' name='"+tx&ts&"c"+"' value='"+rs("choicec")+"' size='100'>"
  dd="<input type='text' onChange='js("+chr(34)+na1+chr(34)+")' name='"+tx&ts&"d"+"' value='"+rs("choiced")+"' size='100'>"
 
 %>
 <tr align="left"><td >A:<%=aa%></td></tr>
 <tr align="left"><td >B:<%=bb%></td></tr>
  <tr align="left"><td >C:<%=cc%></td></tr>
  <tr align="left"><td >D:<%=dd%></td></tr>
  <%end if
  %>
                                                              
 
  
  
<%rs.movenext%><%loop%>            
</table>
<p><input type="submit" name="submit" value="保存试题"></p>

</form>

<%rs.close
  %>
<br><br><center><a href=ask.asp?dh=<%=dh%>&sm=<%=sm%>&key=<%=key%>
>返回到上一页</a> </center>
 <BR>                           
<center><br>　 <a href=index.htm>回到首页</a>                                     
</center>                                   
</body>                                                  
                                                  
</html>