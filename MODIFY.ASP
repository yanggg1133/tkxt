<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>考试号码的浏览</title>
</head>

<body>

<p align="center"><font size="5" face="隶书">选择待评试卷的考试号码</font></p>
<p align="center"><font size="3" face="方正舒体">你必须先选择已经存在的考试号码才能评卷</font></p>

<%                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")                    
rs.open "select distinct kshm,dh,sm,kssj from sj",conn,3,3                    
%>
<table align="center" width="700" border="1" cellspacing="0" cellpadding="0" height="47">
<tr><td witdh="25%" bgcolor="#0000FF" height="16">
    <p align="center"><font color="#FFFFFF">试卷代号</font></p>
  </td><td width="50%" bgcolor="#0000FF" height="16">
    <p align="center"><font color="#FFFFFF">标题</font></p>
  </td><td width="25%" bgcolor="#0000FF" height="16">
    <p align="center"><font color="#FFFFFF">考试时间</font></p>
  </td></tr>
<%do while not rs.eof %>      
<tr onmouseout="this.bgColor='';" onmouseover="this.bgColor='#CFFEDF';">
<td align="center" height="27">
 <%=rs("kshm")%>           
  </td>
<td align="center" height="27">
<a href=modify1.asp?kshm=<%=rs("kshm")%>>,&nbsp<%=rs("sm")%>&nbsp;<%=rs("kssj")%></a>            
  </td><td align="center" height="27">
<%=rs("kssj")%>分钟
  </td></tr>
<%rs.movenext%><%loop%>
<%rs.close
conn.close%>
</table>
<br><p><center><a href=kh.asp>直接返回</a></center></p>
</body>

</html>












