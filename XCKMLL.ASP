<!-- #include file=conn.asp -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 4</title>
</head>

<body>
<% dim exec
exec="select distinct sjdh.sm,sj.kshm from sjdh,sj  where( sj.dh in (select dh from sjdh) and sj.dh=sjdh.dh)"
'set conn=server.createobject("adodb.connection") 
'conn.open "dsn=xiong1" 
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3

%>
<p align="center"><font size="4" face="隶书" color="#000000"><span style="background-color: #FF0066">选择考试科目菜单</span></font></p>
<div align="center">
  <center>
<table border="1" width="700" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#66FF66" height="52" >

<tr><td width="50%" height="22" bgcolor="#0000FF">
<p align="center">
<font size="3" face="华文琥珀" color="#FFFFFF"><b>科目名</b></font></p>
  </td><td height="22" bgcolor="#0000FF">
<p align="center">
<font size="3" face="华文琥珀" color="#FFFFFF"><b> 

考试号码</b></font> </p>
 </td></tr>


<%do while not rs.eof%>              
<tr onmouseout="this.bgColor='';" onmouseover="this.bgColor='#CFFEDF';"><td height="26">
<p align="center">
<a href=cxxt.asp?sm=<%=rs("sm")%>&kshm=<%=rs("kshm")%>><%=rs("sm")%>;</a>              
</p>
</td><td height="26">
<p align="center"><%=rs("kshm")%></p>
</td></tr>
<%rs.movenext%><%loop%>             
<%rs.close
  conn.close
%>
</table>
  </center>
</div>
<p><center><a href=cxxt.asp>直接返回</a></center></p>
</body>

</html>
