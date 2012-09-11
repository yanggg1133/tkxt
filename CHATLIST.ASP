<%refreshtime=10%>

<html>

<head>
<meta http-equiv="refresh" content="<%=refreshtime %>,url=chatlist.asp">
</head>

<body bgcolor="black">
<%
for i=1 to 21
 response.write application("msg"&i)&"<br>"
 next
 %>

</body>

</html>
