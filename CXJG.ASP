<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>
<SCRIPT language=JavaScript>
var currentpos,timer;

function initialize()
{
timer=setInterval("scrollwindow()",80);
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

<%dim sm,kshm,bj,xm,xh,gr,jt,exec,m,all,sum,bt,zd,bjf,gs
sm=session("sm")
kshm=session("kshm")
bj=session("bj")
xm=request.form("name")
xh=request.form("xh")
gr=request.form("gr")
jt=request.form("bj")
all=request.form("all")
%>
<%                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")  
%>                  
<p align="center"><font size="5" face="隶书" color="#FFFF00"><span style="background-color: #FF0000">查 
询 结 果 统 计 表</span></font></p>
<div align="center">
  <center>
  <% if gr<>"" then %>

 <p align="center"><font size="5" face="华文隶书" color="#000066">个人查询结果统计表</font></p>                        
<% end if%>
  <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolordark="#0000FF" bordercolorlight="#008000">
  <% if gr<>"" then %> 
    <tr>
      <td width="8%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>姓名</b></font></td>
      <td width="8%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>学号</b></font></td>
      <td width="12%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>班级</b></font></td>
      <td width="8%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>分数</b></font></td>

      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">选择题</font></td> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">填空题</font></td> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">判断题</font></td> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">计算题</font></td> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">简答题</font></td> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">程序题</font></td> 
    <font color="#000066"> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">完型填空</font></td> 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">程序结果</font></td> 
    </tr> 
 <% end if%>

 <%zd="xscj.xm,xscj.xh,xscj.bj,xscj.xc,xscj.tk,xscj.pd,xscj.js,xscj.jd,xscj.cx,xscj.wxtk,xscj.cxjg,xscj.zf,sjdh.sm"
%>
<% if gr<>"" then 
  if all<>"ON" then 
  m="kshm='"&kshm&"' and" 
  else m="" 
  end if 
if xh<>"" then 
  if xm<>"" then 
  exec= "select "&zd&" from xscj,sjdh where("&m&" xh='"&xh&"' and xscj.gf=1 and xscj.dh=sjdh.dh and xm like  '%"&xm&"%')"  
  else 
 exec= "select "&zd&" from xscj,sjdh where("&m&" xscj.gf=1 and xscj.dh=sjdh.dh and xh='"&xh&"')"  
  end if 
else 
  if xm<>"" then 
   exec= "select "&zd&" from xscj,sjdh where("&m&" xscj.gf=1 and xscj.dh=sjdh.dh and xm like '%"&xm&"%')"  
   else  
   if m<>"" then
   exec= "select "&zd&" from xscj,sjdh where(kshm='"&kshm&"' and xscj.gf=1 and xscj.dh=sjdh.dh"&")" 
   else
   exec="select "&zd&" from xscj,sjdh where(xscj.gf=1 and xscj.dh=sjdh.dh)"
    end if
  end if  
end if 
sum=0
gs=0
bt=""
'response.write exec
'response.end
rs.open exec,conn,3,3%> 
 <%do while not rs.eof%>              
<%if bt<>rs("sm") then 
  bt=rs("sm")%>
  <% if gs>0 and rs.recordcount>1 then %>
  <tr> 
      <td width="27%" colspan="3" height="18"> 
        <p align="center"><span style="background-color: #FFcc00">总&nbsp; 分</span></td>  
      <span style="background-color: #FFFFCC">  
      <td width="8%" height="18"><%=sum%></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
    </tr>  
  <tr>  
      <td width="27%" colspan="3" height="18">  
        <p align="center"><span style="background-color: #FF6600">平均分</span></td>  
      <span style="background-color: #FF6600">  
      <td width="8%" height="18"><%=sum/gs%></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
    </tr> 
     <%end if %> 
  <% gs=0 
     sum=0 %> 
<tr> 
<td width="100%" colspan="12" align="center" height="18"> 
<p align="center"><span style="background-color: #FFFFCC"><font color="#000066"><%=bt%></p> 
</tr> 
<%end if%> 
<tr>  
      <td width="8%" height="18"><%if rs("xm")="" then%><%="***"%><%else%><%=rs("xm")%><% end if%></td>  
      <td width="8%" height="18"><%if rs("xh")="" then%><%="***"%><%else%><%=rs("xh")%><% end if%></td>  
      <td width="12%" height="18"><%if rs("bj")="" then%><%="***"%><%else%><%=rs("bj")%><% end if%></td>  
      <td width="8%" height="18"><%=rs("zf")%></td>  
      <td width="8%" height="18"><%=rs("xc")%></td>  
      <td width="8%" height="18"><%=rs("tk")%></td>  
      <td width="8%" height="18"><%=rs("pd")%></td>  
      <td width="8%" height="18"><%=rs("js")%></td>  
      <td width="8%" height="18"><%=rs("jd")%></td>  
      <td width="8%" height="18"><%=rs("cx")%></td>  
      <td width="8%" height="18"><%=rs("wxtk")%></td>  
      <td width="8%" height="18"><%=rs("cxjg")%></td>  
    </tr>  
<%sum=sum+rs("zf") 
gs=gs+1 
%> 
<%rs.movenext%><%loop%>              
 
          
  <% if rs.recordcount>1 then %> 
    <tr>  
      <td width="27%" colspan="3" height="18">  
        <p align="center"><span style="background-color: #FFcc00">总&nbsp; 分</span></td>  
      <span style="background-color: #FFFFCC">  
      <td width="8%" height="18"><%=sum%></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
    </tr>  
  <tr>  
      <td width="27%" colspan="3" height="18">  
        <p align="center"><span style="background-color: #FF6600">平均分</span></td>  
      <span style="background-color: #FF6600">  
      <td width="8%" height="18"><%=sum/gs%></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
    </tr> 
      
    <%end if%> 
<% rs.close 
 end if 
 %> 
  
 <%bjf=request.form("bj")%> 
 <%if bjf<>"" then %> 
<tr> 
<td width="100%" colspan="12" align="center" height="20"> 
<font size="5" face="华文隶书" color="#ff0066">班级查询结果统计表</font> 
</td></tr> 
 
    <tr> 
      <td width="8%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>姓名</b></font></td> 
      <td width="8%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>学号</b></font></td> 
      <td width="12%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>班级</b></font></td> 
      <td width="8%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>分数</b></font></td> 
 
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">选择题</font></td>  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">填空题</font></td>  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">判断题</font></td>  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">计算题</font></td>  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">简答题</font></td>  
    </font>  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">程序题</font></td>  
    <font color="#000066">  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">完型填空</font></td>  
      <td width="8%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">程序结果</font></td>  
    </tr>  
  
 
<% end if%> 
 
 <%if bjf="ON" then%> 
 <% dim bjname,c1,c2,c3,c4,c5,c6,fs,fw 
 c1=0 
 c2=0 
 c3=0 
 c4=0 
fw=0 
sum=0 
 c5=cint(request.form("start")) 
 c6=cint(request.form("end")) 
 bjname=request.form("bjname") 
 exec="select * from xscj where(kshm='"+kshm+"' and gf=1 and bj='"+bjname+"')" 
 rs.open exec,conn,3,3  
 %> 
 <tr> 
<td width="100%" colspan="12" align="center" height="18"> 
<p align="center"><span style="background-color: #FFFFCC"><font color="#000066"><%=session("sm")%></p> 
</tr> 
 
<%do while not rs.eof%> 
<% fs=rs("zf") 
sum=sum+fs 
if fs<30 then c1=c1+1  
if fs>=30 and fs<60  then c2=c2+1 
if fs>=60 and fs<80 then c3=c3+1 
if fs>=80 and fs<=100 then c4=c4+1 
if fs>=c5 and fs<=c6 then fw=fw+1  
%>  
<tr>  
      <td width="8%" height="18"><%if rs("xm")="" then%><%="***"%><%else%><%=rs("xm")%><% end if%></td>  
      <td width="8%" height="18"><%if rs("xh")="" then%><%="***"%><%else%><%=rs("xh")%><% end if%></td>  
      <td width="12%" height="18"><%if rs("bj")="" then%><%="***"%><%else%><%=rs("bj")%><% end if%></td>  
      <td width="8%" height="18"><%=rs("zf")%></td>  
      <td width="8%" height="18"><%=rs("xc")%></td>  
      <td width="8%" height="18"><%=rs("tk")%></td>  
      <td width="8%" height="18"><%=rs("pd")%></td>  
      <td width="8%" height="18"><%=rs("js")%></td>  
      <td width="8%" height="18"><%=rs("jd")%></td>  
      <td width="8%" height="18"><%=rs("cx")%></td>  
      <td width="8%" height="18"><%=rs("wxtk")%></td>  
      <td width="8%" height="18"><%=rs("cxjg")%></td>  
    </tr>  
 
 <%rs.movenext%><%loop%> 
<% if rs.recordcount>1 then %> 
 
 <tr>  
      <td width="27%" colspan="3" height="18">  
        <p align="center"><span style="background-color: #FFcc00">总&nbsp; 分</span></td>  
      <span style="background-color: #FFFFCC">  
      <td width="8%" height="18"><%=sum%></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
    </tr> 
     <tr>  
      <td width="27%" colspan="3" height="18">  
        <p align="center"><span style="background-color: #FF6600">单科平均</span></td>  
      <span style="background-color: #FFFFCC">  
      <td width="8%" height="18"><%=sum/rs.recordcount%></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
      <td width="8%" height="18"></td>  
    </tr> 
<% end if%> 
 <tr> 
<td width="100%" colspan="12" align="center" height="20"> 
<font size="4" face="华文隶书" color="#000066">分数段统计表</font> 
</td></tr> 
 <% DIM cc1,cc2,cc3,cc4 
 cc1=request.form("cc1") 
 cc2=request.form("cc2") 
 cc3=request.form("cc3") 
 cc4=request.form("cc4") 
%> 
 <%if cc1="on" or cc2="on" or cc3="on" or cc4="on" then %> 
 
<tr> 
 
 <%if cc1="on" then%> 
 <td colspan="2"  height="18">30分以下</td><td colsapn="2"><%=c1%>　</td> 
 <%end if%> 
 <% if cc2="on" then%> 
 <td colspan="2"  height="18">30-59</td><td ><%=c2%>　</td> 
 <%end if%> 
 <% if cc3="on" then%> 
 <td colspan="2"  height="18">60-79</td><td colsapn="2"><%=c3%>　</td> 
 <%end if%> 
 <% if cc4="on" then%> 
 <td colspan="2"  height="18">80-100</td><td colsapn="2"><%=c4%>　</td> 
 <%end if%>  
  
 </tr> 
 <%end if%> 
 <%if c6<>-1 and c6>=c5 and c5>=0  then%> 
 
  <tr> 
<td width="100%" colspan="12" align="center" height="20"> 
<font size="4" face="华文隶书" color="#000066">指定分数段统计人数表</font> 
</td></tr> 
 
 <tr> 
  <td colspan="2"  height="18"><%=c5%>--<%=c6%>有：</td><td colsapn="10"><%=fw%>　</td> 
</tr> 
 <%end if%> 
  
  <%rs.close%> 
 
 <%end if%> 
  
  
 </table>  
  
<p align="center"><a href="cxxt.asp">返回到查询页面&nbsp;</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="index.htm"> 
返回首页</a></p>      
 
</body> 
 

</html>