<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>考试号码的浏览</title>
</head>

<body>

<p align="center"><font size="5" face="隶书">选择考试该试卷的考生</font></p>
<p align="center"><font size="3" face="隶书"><%=session("sm")%>=<%=session("kshm")%></font></p>

<%                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")                    
rs.open "select * from xscj where(kshm='"&session("kshm")&"')",conn,3,3                    
%>
<table align="center" width="100%" height="54" border="1" cellspacing="0" cellpadding="0">
<tr><td witdh="25%" width="184" height="18" bgcolor="#0000FF">
    <p align="center"><font color="#FFFFFF">学  号</font></p> 
  </td><td width="25%" height="18" bgcolor="#0000FF"> 
    <p align="center"><font color="#FFFFFF">姓  名</font></p> 
  </td><td width="25%" height="18" bgcolor="#0000FF"> 
    <p align="center"><font color="#FFFFFF">班  级</font></p> 
  </td>
  <td width="25%" height="18" bgcolor="#0000FF"> 
    <p align="center"><font color="#FFFFFF">删除否</font></p> 
  </td>
  </tr> 
<%do while not rs.eof %>       
<tr onmouseout="this.bgColor='';" onmouseover="this.bgColor='#CFFEDF';">
<td align="center" width="184" height="28">
<%=rs("xh")%>             
</td> 
<td align="center" width="319" height="28"> 
<a href=modify3.asp?xh=<%=rs("xh")%>> 
<%if rs("gf")=1 then %>
（已评）<span style="background-color: #66FF66"><font color="#FF0000" size="3" face="隶书"> 
<%end if%> 
<%=rs("xm")%>
<%if rs("gf")=1 then %>  
   </font></span> 
   <%end if%> 
</a>               
  </td><td align="center" width="239" height="28"> 
   <%=rs("bj")%>
  </td> 
   </td><td align="center" width="239" height="28"> 
<a href=dele.asp?xh=<%=rs("xh")%>&kshm=<%=session("kshm")%>>删除</a>
  </td>
  </tr>
<%rs.movenext%><%loop%> 
</table> 
<br><p><center><a href=kh.asp>直接返回</a></center></p> 
</body> 
 
</html>