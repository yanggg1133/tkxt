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

<%dim hm,value,pp,exec,sum,gs,zd,hm1
  dim zz(5)
  zz(1)="姓名"
  zz(2)="准考证号"
  zz(3)="班级"
  zz(4)="报名号"
  zz(5)="专业"
hm=request.form("item")
hm1=zz(cINT(hm))
value=request.form("va")
%>
<%                    
set rs=server.createobject("adodb.recordset")  
%>                  
<p align="center"><font size="5" face="隶书" color="#FFFF00"><span style="background-color: #FF0000">查 
询 结 果 统 计 表</span></font></p>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolordark="#0000FF" bordercolorlight="#008000">
 
    <tr>
      <td width="14%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>准考证号</b></font></td>
      <td width="12%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>姓名</b></font></td>
      <td width="12%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>班级</b></font></td>
      <td width="12%" bgcolor="#808080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>专业</b></font></td>
    
      <td width="5%" bgcolor="#ff8080" align="center" height="18"><font size="3" face="楷体_GB2312" color="#000099"><b>总分</b></font></td>

      <td width="5%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">语文</font></td> 
      <td width="5%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">数学</font></td> 
      <td width="5%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">英语</font></td> 
      <td width="5%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">综合</font></td> 
      <td width="5%" bgcolor="#C0C0C0" align="center" height="18"><font size="2" color="#0066FF">名次</font></td> 
      </tr> 
 

 <%zd="gkcj.准考证号,gkcj.姓名,gkcj.班级,gkcj.专业,gkcj.总分,gkcj.语文,gkcj.数学,gkcj.英语,gkcj.综合,gkcj.名次"
%>
<%
if value<>"" then 
  if value="*" then
  exec= "select "&zd&" from gkcj"
  else
  exec= "select "&zd&" from gkcj where("&hm1&"='"&value&"')"  
  end if

  else 
   response.end 
  end if 
'response.write exec
'response.end

sum=0
gs=0
rs.open exec,conn%> 
 <%do while not rs.eof%>              
<tr>  
      <td width="14%" height="18"><%if rs("准考证号")="" then%><%="***"%><%else%><%=rs("准考证号")%><% end if%></td>  
      <td width="12%" height="18"><%if rs("姓名")="" then%><%="***"%><%else%><%=rs("姓名")%><% end if%></td>  
      <td width="12%" height="18"><%if rs("班级")="" then%><%="***"%><%else%><%=rs("班级")%><% end if%></td>  
      <td width="12%" height="18"><%if rs("专业")="" then%><%="***"%><%else%><%=rs("专业")%><% end if%></td>
      <td width="12%" height="18"><%=rs("总分")%></td>  
      <td width="5%" height="18"><%=rs("语文")%></td>  
      <td width="5%" height="18"><%=rs("数学")%></td>  
      <td width="5%" height="18"><%=rs("英语")%></td>  
       <td width="5%" height="18"><%=rs("综合")%></td>  
      <td width="5%" height="18"><%if rs("名次")="" then%><%="***"%><%else%><%=rs("名次")%><% end if%></td>  
 </tr>  
<%sum=sum+rs("总分") 
gs=gs+1 
%> 
<%rs.movenext%><%loop%>              
 
          
  <% if gs>1 then %> 
    <tr>  
      <td width="27%" colspan="4" height="18">  
        <p align="center"><span style="background-color: #FFcc00">总&nbsp; 分</span></td>  
      <span style="background-color: #FFFFCC">  
      <td  height="18"><%=sum%></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
       
    </tr>  
  <tr>  
      <td width="27%" colspan="4" height="18">  
        <p align="center"><span style="background-color: #FF6600">平均分</span></td>  
      <span style="background-color: #FF6600">  
      <td  height="18"><%=sum/gs%></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
      <td  height="18"></td>  
  
   </tr> 
      
    <%end if%> 
<% rs.close 
 %> 
  
 
 </table>  
  
</div>
 
</body> 
 </html>