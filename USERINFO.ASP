<%
if request("logout")="yes" then
application.lock
for i=21 to 2 step -1
j=i-1
application("msg"&i)=application("msg"&j)
next
strhour=hour(time())
if len(strhour)=1 then strhour="0"&strhour
strminute=minute(time())
if len(strminute)=1 then strminute="0"&strminute
strtime="<"&strhour&":"&strminute&">&nbsp:"
strtmp="<font color='"&session("usercolor")&"'>"&"<img src='"&session("userfig")&"'>"&session("username")&"说:我要离开聊天室了!,!!"&strtime&"</font>"
application.unlock


%>
<script language="vbscript">
parent.window.close
</script>
<%end if%>
<html>


<body bgcolor="namy" text="white">
<img src="1.gif" align="left">
<%=session("username")%>于<%=now()%>进入了江南职高聊天室.ip为<%=request.servervariables("REMOTE_HOST")%>
<a href="userinfo.asp?logout=yes"><img src="2.gif"  border="0" align="center"></a>


</body>

</html>
