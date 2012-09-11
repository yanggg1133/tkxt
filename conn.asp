<%@ CODEPAGE = "936" %>
<%'on error resume next
 feiyudbq="kkk.mdb"'如果你更改了数据库文件名称，请更改此处
   connstr="DBQ="+server.mappath(""&feiyudbq&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
     set conn=server.createobject("ADODB.CONNECTION")
     conn.open connstr 
response.buffer=true
Response.ExpiresAbsolute = Now() - 1  
Response.Expires = 0  
Response.CacheControl = "no-cache"
%>