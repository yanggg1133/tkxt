<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���Ժ�������</title>
</head>

<body>

<p align="center"><font size="5" face="����">ѡ���Ը��Ծ��Ŀ���</font></p>
<p align="center"><font size="3" face="����"><%=session("sm")%>=<%=session("kshm")%></font></p>

<%                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")                    
rs.open "select * from xscj where(kshm='"&session("kshm")&"')",conn,3,3                    
%>
<table align="center" width="100%" height="54" border="1" cellspacing="0" cellpadding="0">
<tr><td witdh="25%" width="184" height="18" bgcolor="#0000FF">
    <p align="center"><font color="#FFFFFF">ѧ  ��</font></p> 
  </td><td width="25%" height="18" bgcolor="#0000FF"> 
    <p align="center"><font color="#FFFFFF">��  ��</font></p> 
  </td><td width="25%" height="18" bgcolor="#0000FF"> 
    <p align="center"><font color="#FFFFFF">��  ��</font></p> 
  </td>
  <td width="25%" height="18" bgcolor="#0000FF"> 
    <p align="center"><font color="#FFFFFF">ɾ����</font></p> 
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
��������<span style="background-color: #66FF66"><font color="#FF0000" size="3" face="����"> 
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
<a href=dele.asp?xh=<%=rs("xh")%>&kshm=<%=session("kshm")%>>ɾ��</a>
  </td>
  </tr>
<%rs.movenext%><%loop%> 
</table> 
<br><p><center><a href=kh.asp>ֱ�ӷ���</a></center></p> 
</body> 
 
</html>