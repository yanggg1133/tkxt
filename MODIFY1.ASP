<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���Ժ�������</title>
</head>

<body>

<p align="center"><font size="5" face="����">ѡ������Ծ��Ŀ��Ժ���</font></p>
<p align="center"><font size="3" face="��������">�������ѡ���Ѿ����ڵĿ��Ժ����������</font></p>

<% dim kshm
kshm=request.querystring("kshm")                    
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset")                    
rs.open "select * from sj where(kshm='"&kshm&"')",conn,3,3                    
  if not rs.eof then     
   session("exec")=rs("tx")   '��������              
session("dh")=rs("dh")      '�Ծ�����           
session("sm")=rs("sm")      '�Ծ�����           
session("kssj")=rs("kssj")   '����ʱ��             
session("kshm")=rs("kshm")  '���Ժ���         
response.redirect "modify2.asp"  
 else%> 
<center>�����ڸ��Ծ�,�뷵��....</center><p>
<a href="kh.asp"><span style="background-color: #FF6666"><font size="4" face="��������" color="#66FF33">���ص���һҳ</font></span></a>
 </p>

<%end if
rs.close
conn.close%>
</body>

</html>









