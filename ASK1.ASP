<%@language=vbscript %>
<%
dim que,dif 
 
 que=session("section")
dif= session("diff")  



 %> 
 <script language="javascript">
function check(myform)
{
if(myform.question.value=="")
{alert("请输入题目");
 return false;
}
if(myform.choicea.value=="")
{alert("请输入选项A");
return false;
}
if(myform.choiceb.value=="")
{alert("请输入选项B");
return false;
}
if(myform.choicec.value=="")
{alert("请输入选项c");
return false;
}
if(myform.choiced.value=="")
{alert("请输入选项d");
return false;
}
}
</script>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>


<body bgcolor="#ffffff">
<form method="post" action="askinto.asp"  name="myform" onsubmit ="return check(this)">
<table border="0">
 <tr>
 <td align="center" valign="top"><p>请输入问题:</td>
            
      
<td align="center" valign="bottom">
&nbsp;<textarea  name="question" cols="80" rows="6"></textarea> </td>
</tr>
</table>
<%if que="选择题" then %>
<p>请输入A选项:                               
<input type="text" name="choicea" size="80" maxlength="200">                               
100字内</p>                               
<p>请输入b选项:                               
<input type="text" name="choiceb" size="80" maxlength="200">                               
100字内</p>                               
<p>请输入c选项:                               
<input type="text" name="choicec" size="80" maxlength="200">                               
100字内</p>                               
<p>请输入d选项:                               
<input type="text" name="choiced" size="80" maxlength="200">                               
100字内</p>              
<%end if%>                              
<p>答案是:                               
<%if que="选择题" then %>&nbsp;&nbsp;&nbsp;                          
<input type="radio" name="answer" value="A" checked>                               
A                               
<input type="radio" name="answer" value="B" >                               
B                               
<input type="radio" name="answer" value="C" >                               
C                               
<input type="radio" name="answer" value="D"  >                               
D                               
<%else if que="判断题" then%>                 
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                 
        <input type="radio" name="answer" value="正确" checked>                               
正确                               
<input type="radio" name="answer" value="错误" >                               
错误                  
<%else%>               
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<textarea  name="answer" cols="80" rows="3"></textarea>                     
               
<%end if%>                
<%end if%>                         
</P>                          
                           
<P> <center>                              
<input  type="submit" name="submit" value="填好了">                               
                               
<input  type="reset" name="submit2" value="重填一次">                               
</certer>                 
</p>        
<center>        
        
<p>        
<a  href=ask.asp?dh=<%=session("dh")%>&sm=<%=session("sm")%>&key=<%=session("key")%>>返回到上一页，继续选择题型和难度</a>        
</p>                               
<p><a href=index.htm>回到首页</a></p>       
</form>                            
</body>                           
                           
</html>