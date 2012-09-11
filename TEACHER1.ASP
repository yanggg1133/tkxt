<%@language=vbscript %>

<%
dim dh,sm,key,tk,xc,pd,js,jd,cx,wxtk,cxjg
tk=session("tk")
xc=session("xc")
pd=session("pd")
js=session("js")
jd=session("jd")
cx=session("cx")
wxtk=session("wxtk")
cxjg=session("cxjg")
dh=request.querystring("dh")
if dh=""  then dh=session("dh")
sm=request.querystring("sm")
if sm="" then sm=session("sm")
key=request.querystring("key")
if key="" then key=session("key")
session("dh")=dh
session("sm")=sm
session("key")=key
%>

<script language="vbscript">
sub button1_onclick
 msgbox "可以通过浏览来选择已有的试卷代号和说明以及关键词"
 end sub
 sub button2_onclick
 msgbox "最好通过浏览本套题存在的题型和数量"
 end sub
  sub button3_onclick
 msgbox "通过浏览可选择一套已经存在的试卷来考试,而不必进行前面的设置"
 end sub

 
 </script> 
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
<script language="javascript">
function check(form1)
{if(form1.sjdh.value=="")
{alert("请输入试卷代号");
return false;
}
}
</script>
</head>

<body bgcolor="#ffffff">


<p align="center"><font size="6" face="隶书" color="#FF00FF"><span style="background-color: #FF0000"><!--欢迎进入选择菜单--><img src="aaa.gif"></span></font></p>
<!--<form action="teacher2.asp" method="post" name="form1" onsubmit ="return check(this)"> -->
<form action="teacher2.asp" method="post" name="form1">

<center><font color="#0000FF" size="4" face="楷体_GB2312"><span style="background-color: #FFCCFF">选择试卷代号：</span></font>
<input type="text" name="sjdh" size="10" maxlength="10" value='<%=dh%>'> &nbsp;&nbsp;&nbsp;   
<input type="button" name="button1" value="提示" >----><a href=teacherll.asp?asp=teacher1>浏览</a></center>   
<%if sm<>"" then%>   
 说明是：<%=sm%>  
<% end if%>   
<%if key<>"" then%>   
<p>关键词：   
<%=key%> </p>  
<%end if%>   
<p>请选择难度:     
<input type="checkbox" name="diff1" value="hard" >     
较难     
<input type="checkbox" name="diff2" value="normal" checked>     

中等     
<input type="checkbox" name="diff3" value="easy" >     
容易&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;   
<input type="button" name="button2" value="提示" >---->
浏览否?<input type="radio" name="sf" value="yes" >是     
<input type="radio" name="sf" value="no" checked>否  
 <input type="submit" name="submit0" value="浏         览"></p>
<p>请选择题型:     
<input type="checkbox" name="xc" value="选择题" <%if xc<>"" then %> <%="checked"%> <%end if%>>选择题 
 <input type="checkbox" name="tk" value="填空题" <%if tk<>"" then %> <%="checked"%> <%end if%>>填空题 
<input type="checkbox" name="pd" value="判断题" <%if pd<>"" then %> <%="checked"%> <%end if%>>判断题 
<input type="checkbox" name="js" value="计算题" <%if js<>"" then %> <%="checked"%> <%end if%>>计算题  
<input type="checkbox" name="jd" value="简答题" <%if jd<>"" then %> <%="checked"%> <%end if%>>简答题 
<input type="checkbox" name="cx" value="程序题" <%if cx<>"" then %> <%="checked"%> <%end if%>>程序题    
   <input type="checkbox" name="wxtk" value="完型填空" <%if wxtk<>"" then %> <%="checked"%> <%end if%>>完型填空 
   <input type="checkbox" name="cxjg" value="程序结果" <%if cxjg<>"" then %> <%="checked"%> <%end if%>>程序结果 
<p><table width="720" border="1">
  <tr align="center">
  <td width="80">题的类型</td>
   <td witdh="11">选择题</td>
   <td witdh="11%">填空题</td>
   <td witdh="11%">判断题</td>
   <td witdh="11%">计算题</td>
   <td witdh="11%">简答题</td>
   <td witdh="11%">程序题</td>
   <td witdh="11%">完型填空</td>
   <td witdh="11%">程序结果</td>
 </tr>  
  <tr align="center">
   <% if xc="" then xc=0
      if tk="" then tk=0
      if pd="" then pd=0
      if js="" then js=0
      if jd="" then jd=0
      if cx="" then cx=0
      if wxtk="" then wxtk=0
      if cxjg="" then cxjg=0
  %>          
   <td width="80">题的数量</td>
   <td witdh="11"><%=xc%></td>
   <td witdh="11%"><%=tk%></td>
   <td witdh="11%"><%=pd%></td>
   <td witdh="11%"><%=js%></td>
   <td witdh="11%"><%=jd%></td>
   <td witdh="11%"><%=cx%></td>
   <td witdh="11%"><%=wxtk%></td>
   <td witdh="11%"><%=cxjg%></td>
   
   </tr>
   </table>
<center>
<P>     
<%
kshm1=request.querystring("kshm")
if kshm1="" then
kshm1=cstr(year(date()))+cstr(month(date()))+cstr(day(date()))
 randomize
 kshm1=kshm1&cstr(fix((100*rnd())))
 end if
 %>
请选择?<input type="radio" name="sf1" value="no" checked> 阅读试卷    
<input type="radio" name="sf1" value="yes" >确定为考试试卷 ===>>
考试时间<input type="text" name="kssj" value=100  size="3" maxlength="3">分 &nbsp;&nbsp;

<input type="radio" name="sf1" value="yesno" > 取消试卷 &nbsp;&nbsp; &nbsp;
<input type="submit" name="submit1" value="下一步"> &nbsp;&nbsp;&nbsp;
</p> 考试号码为:<input type="text" name="kshm" value=<%=kshm1%>  size="10" maxlength="10">
<input type="button" name="button3" value="提示" >----><a href=kshmll.asp>浏览</a>   

</center>

</form>

<center>
  <p><a href="qx.asp">取消试卷</a> <a href=index.htm>回到首页</a></p>
</center>

</body>   
   
</html>   





























































































































