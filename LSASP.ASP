<%@language=vbscript%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>

<body>
<%dim k,d,x,m
x=2
k="d"&x
m="<input type='text' name='"+k+"' size='10' maxlength='10'>"
%>

<form method="post" action="aspls.asp"  name="myform" >
<p>输入试卷代号： </p><%=m%>
<input type="submit" name="submit" value="填好了">    

</form> 

</body>

</html>
