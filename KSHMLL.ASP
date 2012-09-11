<!--#include file="conn.asp" -->
<html>
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

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>考试号码的浏览</title>
</head>

<body>

<p align="center"><font size="5" face="隶书">考试号码的浏览</font></p>
<p align="center"><font size="3" face="方正舒体">你可以通过浏览来选择已经存在的考试号码</font></p>

<%                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")                    
rs.open "select distinct sj.kshm,dh,sm,kssj from sj",conn,3,3                    
%>
<table align="center" width="100%">
<tr><td witdh="25%">试卷代号</td><td width="50%">标题(考试号码)</td><td width="25%">考试时间</td></tr>
<%do while not rs.eof %>      
</tr><td align="center" colspan="3">
<a href=teacher1.asp?kshm=<%=rs("kshm")%>><%=rs("dh")%>,&nbsp<%=rs("sm")%>&nbsp;(<%=rs("kshm")%>)<%=rs("kssj")%><br></a>            
  </td></tr>
<%rs.movenext%><%loop%>
</table>
<br><p><center><a href=teacher1.asp?kshm="">直接返回</a></center></p>
</body>

</html>











