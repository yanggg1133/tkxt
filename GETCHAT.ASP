<%
application.lock
for i=21 to 2 step -1
j=i-1
application("msg"&i)=application("msg"&j)
next
strhour=hour(time())
if len(strhour)=1 then strhour="0"&strhour
strminute=minute(time())
if len(strminute)=1 then strminute="0"&strminute
strtime="<"&strhour&":"&strminute&">"
strtmp="<font color='"&session("usercolor")&"'>"&"<img src='"&session("userfig")&"'>"&session("username")&"˵:"&request.form("chatmsg") &strtime&"</font>"
application("msg1")=strtmp
application.unlock
response.redirect "chatlist.asp"
%>
