<!--#include file="conn.asp" -->
<%  
dim xh,kshm 
xh=request.querystring("xh")
kshm=request.querystring("kshm")
exec="delete from da  where xh='"+xh+"' and kshm='"+kshm+"'"   
exec="delete from xscj  where xh='"+xh+"' and kshm='"+kshm+"'"   
conn.execute exec   

response.write "<br><br><br><br><br><center><b>已成功地删除该生的答案</b>"   
    %> 
 <html>
<head>
<meta http-equiv="refresh" content="4;url=modify2.asp?xh=<%=xh%>&kshm=<%=kshm%>">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 4</title>
</head>
<body>
<p>
<a href=modify2.asp?xh=<%=xh%>&kshm=<%=kshm%>>返回上一页</a>
</p>
</body>     
 </html>