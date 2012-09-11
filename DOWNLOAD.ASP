<%
function counts(counterfile)
dim objfso,objts
application.lock
set objfso=server.createobject("scripting.filesystemobject")
set objts=objfso.opentextfile(server.mappath(counterfile),1,true)
if not objts.atendofstream then
 counts=clng(objts.readline)
 end if
 counts=counts+1
 objts.close
 set objts=objfso.opentextfile(server.mappath(counterfile),2,true)
 objts.writeline(counts)
 objts.close
 application.unlock
 end function
 
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>

<body bgcolor="#3366FF">

<p align="center">　</p>
<p align="center"><span style="background-color: #6666FF"><font color="#FF0000" size="6" face="华文琥珀">欢迎下载本网站的系统文件</font></span></p>
<p align="center"><span style="background-color: #FFFF00">你是第位<%=counts("counter.txt")%>访问者</span></p>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolor="#66FF33">
    <tr>
      <td width="33%" height="30">
        <p align="center"><font color="#FFFF99">文件名</font></td>
      <td width="33%" height="30">
        <p align="center"><font color="#FFFF99">文件说明</font></td>
      <td width="34%" height="30">
        <p align="center"><font color="#FFFF99">下载文件</font></td>
    </tr>
    <tr>
      <td width="33%" height="30">
        <p align="center">index.htm</td>
      <td width="33%" height="30">
        <p align="center">系统主文件</td>
      <td width="34%" height="30">
        <p align="center"><a href="counter.exe">下载</a></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
    <tr>
      <td width="33%" height="30"></td>
      <td width="33%" height="30"></td>
      <td width="34%" height="30"></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
