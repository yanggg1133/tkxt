<!--#INCLUDE FILE="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>sdfasdf</title>
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

dim exec,dh,sm,key,m,kshm,tx,ts,sf,sfa,sfb,sfc,sfd
function zh(fString)

    fString = replace(fString , "&gt;", ">")
    fString = replace(fString, "&lt;", "<")
    fString = replace(fString, "&nbsp;", " ")
    
    fString = Replace(fString, "<BR>", CHR(13)+chr(10))
    
    zh = fString
end function

sm=session("sm")
dh=session("dh")  
key=session("key")  
exec="select * from question where(dh='"+dh+"')"+"order by section desc"
session("exec")=exec
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,3,3
%>
<center><span style="background-color: #FFFF00"><font face="楷体_GB2312" color="#0000FF" size="5"><%=session("sm")
%></font></span>
<form method="post" action="da1.asp"  name="myform"> 
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
  <%end if
  sf=rs("answer")%>
  <%  na=tx&ts %>                                                               
  <tr><td width="100%"  colspan="2">                                                               
<% if tx="选择题" then  
  
  if sf="A" or sf="a" then sfa="checked" else sfa=""
  if sf="B" or sf="b" then sfb="checked" else sfb=""
  if sf="C" or sf="c" then sfc="checked" else sfc=""
  if sf="D" or sf="d" then sfd="checked" else sfd=""
                                                             
   m="请选择：<input type='radio' name='"+na+"' value='A'"+sfa+"> A"%>                                                               
  <%=m%>&nbsp;&nbsp;&nbsp;&nbsp;                                                               
 <% m="<input type='radio' name='"+na+"' value='B'"+sfb+"> B"%>                                                               
<%  =m%> &nbsp;&nbsp;&nbsp;&nbsp;                                                              
  <%m="<input type='radio' name='"+na+"' value='C'"+sfc+"> C"%>                                                               
 <% =m     %>&nbsp;&nbsp;&nbsp;&nbsp;                                                               
<%m="<input type='radio' name='"+na+"' value='D'"+sfd+"> D"%>                                                               
 <% =m %>&nbsp;&nbsp;&nbsp;&nbsp;                                                               
  <%else                                                                
  if tx="判断题" then 
    if sf="yes" then sfa="checked" else sfa=""
    if sf="no" then sfb="checked" else sfb=""                                                          
  m="正确否：<input type='radio' name='"+na+"' value='yes'"+sfa+">正确"%>                                                               
  <%=m%>&nbsp;&nbsp;&nbsp;&nbsp;                                                               
 <% m="<input type='radio' name='"+na+"' value='no'"+sfb+">错误"%>                                                               
<%  =m%>                                                               
                                                                 
<%else                                                             
  if tx="填空题"  then
  
  m="答案是：<textarea  name='"+na+"' cols='94' rows='2'>"+zh(sf)+"</textarea>"
  else
  m="答案是：<textarea  name='"+na+"' cols='94' rows='6'>"+zh(sf)+"</textarea>"
  %>
  
   <%end if%>                                                               
   <%=m%>                                                              
  <%end if%>                                                               
  <%end if%></td></tr> 
  
  
<%rs.movenext%><%loop%>            
</table>
<p><input type="submit" name="submit" value="保存答案"></p>

</form>
<%rs.close
  %>
<br><br><center><a href=ask.asp?dh=<%=dh%>&sm=<%=sm%>&key=<%=key%>
>返回到上一页</a> </center>
                               
<center><br>　 <a href=index.htm>回到首页</a>                                     
</center>                                   
</body>                                                  
                                                  
</html>