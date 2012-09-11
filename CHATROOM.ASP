<%
session("username")=request.form("username")
session("usercolor")=request.form("usercolor")
session("userfig")=request.form("userfig")
if session("username")="" then
response.write "很抱歉,您尚未登录,无法进入聊天室!"
response.end
end if
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
strtmp="<font color='"&session("usercolor")&"'>"&"大家好,我是"&session("username")&",请多多指教!"&strtime&"</font>"
application("msg1")=strtmp
application.unlock
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>江南职高聊天室</title>
</head>

<frameset rows="56,*,101">
<frame name="top" noresize scrolling="no" src="userinfo.asp?logout=no">
<frameset cols="160,*">

<frame name="lmiddle" noresize src="userlist.asp?username=<%=session("username")%>">
<frame name="rmiddle" noresize src="chatlist.asp">
</frameset>
<frame name="bottom" noresize scrolling="no" src="inputmsg.asp">
</frameset>

</html>
