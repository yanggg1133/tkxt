<%
dim refreshtime,idletime,totalusers,onlineuser(),tmp(),num,i,newuser,user,username
refreshtime=8
idletime=15
newuser=request("username")
application.lock
if application(newuser&"lastaccesstime")=empty then
  if application("totalusers")=empty then 
  application("totalusers")=0
  end if
redim tmp(application("totalusers")+1)
'response.write application("totalusers")
num=0
if application("totalusers")>0 then
 for i=lbound(application("onlineuser")) to ubound(application("onlineuser"))
 user=application("onlineuser")(i)
 if user<>newuser and user<>session("username") then
 tmp(num)=user
 num=num+1
 else
 application(user&"lastaccesstime")=empty
 end if
 next
 end if
 session("username")=newuser
 tmp(num)=session("username")
 application("totalusers")=num+1
 redim preserve tmp(application("totalusers"))
 application("onlineuser")=tmp
 end if
 
 
 application(session("username")&"lastaccesstime")=timer
 redim tmp(application("totalusers"))
 num=0
 for i=0 to application("totalusers")-1
 user=application("onlineuser")(i)
 if(timer-application(user&"lastaccesstime"))<idletime then
 tmp(num)=user
 num=num+1
 else
 application(user&"lastaccesstime")=empty
 end if
 next
 
 
 if num<>application("totalusers") then
 redim preserve tmp(num)
 application("onlineuser")=tmp
 application("totalusers")=num
 end if
 application.unlock
 
%>
<html>

<head>
<meta http-equiv="refresh"  Content="<%=refreshtime%>,url=<%=request.servervariables("path_info")%>?username=<%=request("username")%>">

</head>

<body bgcolor="black" text="white">
目前在线人数:<%=application("totalusers")%><br>
<ol type="1">
<% for i=0 to (application("totalusers")-1)
response.write "<li>"&application("onlineuser")(i)&"</li>"
next %>

</ol>
</body>

</html>
