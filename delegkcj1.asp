<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>
<body>

<%dim hm,value,pp,exec,hm1
  dim zz(5)
  zz(1)="����"
  zz(2)="׼��֤��"
  zz(3)="�༶"
  zz(4)="������"
  zz(5)="רҵ"
hm=request.form("item")
hm1=zz(cINT(hm))
value=trim(request.form("va"))
set rs=server.createobject("adodb.recordset")  
%>                  
 

 
<%
if value<>"" then 
  if value="*" then
  exec= "delete  from gkcj"
  else
  exec= "delete  from gkcj where trim("&hm1&")='"&value&"'"  
  end if
  conn.execute exec
  %>
  <p align="center">�ɹ�ɾ��</p>
<br>  
<a align="center" href="delegkcj.asp">����</a>
<%else%>
<p align="center">δɾ���κμ�¼</p>
<br>  <a align="center" href="delegkcj.asp">����</a>
   <%
   response.end 
  end if 
%>

  
</body> 
 </html>