<%dim username,userpass,km,bj,bj1,xh,sjname
sjname=request.querystring("sm")
bj1=request.querystring("bj1")
xh=request.querystring("xh")
km=request.querystring("km")
bj=request.querystring("bj")
 username=request.querystring("username")
 userpass=request.querystring("userpass")
 %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>

<body>
<p align="center"><font face="楷体_GB2312" size="6" color="#FFFF00"><span style="background-color: #FF0066">欢  
迎 进 入 管 理 系 统</span></font></p> 
 <form method="POST" action=command.asp>

<input TYPE="hidden" NAME="VTI-GROUP" VALUE="0">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#FFFF00" bordercolordark="#FFFFFF" height="325">
    <tr>
      <td width="100%" colspan="3" height="40" bgcolor="#666699">
        <p align="center"><font size="5" face="楷体_GB2312" color="#000066"><span style="background-color: #C0C0C0"><b><input type="checkbox" name="user" value="ON">用  
        户 管 理</b></span></font></td> 
    </tr> 
    <tr> 
      <td width="24%" height="40" align="center" bgcolor="#666699"><font size="4" face="华文新魏"><input type="radio" name="user1" value="1">用户浏览</font></td>
    </center>
    <td width="76%" height="40" bgcolor="#C0C0C0" colspan="2">
      
        <p align="center"><input type="text" name="userlist" size="20"></p>
     
    </td>
  </tr>
  <center>
  <tr>
    <td width="24%" height="40" align="center" bgcolor="#666699">
      <p align="center"><font size="4" face="华文新魏"><input type="radio" name="user1" value="2" checked>增加用户</font></td>
    <td width="76%" height="40" bgcolor="#C0C0C0" colspan="2">
      <p align="center">用户名<input type="text" name="useradd" size="20"> 密码<input type="text" name="pass" size="20"></p>
    </td>
  </tr>
  <tr>
    <td width="24%" height="40" align="center" bgcolor="#666699"><font size="4" face="华文新魏"><input type="radio" name="user1" value="3">删除用户</font></td>
    <td width="76%" height="40" bgcolor="#C0C0C0" colspan="2">
      <p align="center"><input type="text" name="userdel" size="20" value= '<%=username%>'>浏览否？<input type="radio" value="yes" name="sfdel" checked>是<input type="radio" name="sfdel" value="no">否 
      <input type="submit" value="浏览" name="B6"></p>
    </td>
  </tr>
  <tr>
    <td width="24%" height="1" align="center" bgcolor="#666699" rowspan="2"><font size="4" face="华文新魏"><input type="radio" name="user1" value="4">修改用户</font></td>
    <td width="20%" height="1" bgcolor="#C0C0C0" rowspan="2">
      <p align="center"><input type="radio" value="yes" name="modill" checked>是<input type="radio" name="modill" value="no">否 
      <input type="submit" value="浏览" name="B2"></td>
    <td width="56%" height="21" bgcolor="#C0C0C0" align="center">用户名
    <input type="text" name="username" size="20" value='<%=username%>'></td>
  </tr>
  <tr>
    <td width="56%" height="1" bgcolor="#C0C0C0" align="center">密&nbsp; 码 
    <input type="text" name="userpass" size="20" value='<%=userpass%>'></td>                         
  </tr>                         
  <tr>                         
    <td width="100%" colspan="3" height="41" align="center" bgcolor="#666699"><font size="5" face="楷体_GB2312"><b><span style="background-color: #C0C0C0"><input type="checkbox" name="tkgl" value="ON">题                           
      库 管 理</span></b></font></td>                          
  </tr>                          
  <tr>                          
    <td width="24%" rowspan="2" height="62" align="center" bgcolor="#666699"><font size="4" face="华文新魏"><input type="checkbox" name="tkgl1" value="ON">删除学生答卷</font></td>                         
    <td width="36%" height="31" bgcolor="#C0C0C0">科目<input type="text" name="km" size="20" value="<%=km%>"></td>                         
    <td width="40%" height="62" bgcolor="#C0C0C0" rowspan="2">                         
      <p align="center"><input type="radio" value="yes" name="kmll" checked>是<input type="radio" name="kmll" value="no">否                          
      <input type="submit" value="浏览" name="B3"></td>                         
  </tr>                         
  <tr>                         
    <td width="36%" height="31" bgcolor="#C0C0C0">班级<input type="text" name="bj" size="20" value="<%=bj%>"></td>                         
  </tr>                         
  <tr>                         
    <td width="24%" height="52" align="center" bgcolor="#666699" rowspan="3"><font size="4" face="华文新魏">&nbsp;&nbsp;&nbsp;                          
      <input type="checkbox" name="tkgl2" value="ON">删除学生成绩档案</font></td>                         
    <td width="38%" height="31" bgcolor="#C0C0C0">学号<input type="text" name="xh" size="20" value="<%=xh%>"></td>                         
    <td width="38%" height="28" bgcolor="#C0C0C0" rowspan="2">                         
      <p align="center"><input type="radio" value="yes" name="xhll" checked>是<input type="radio" name="xhll" value="no">否                          
      <input type="submit" value="浏览" name="B4"></p>                         
      <p>　</td>                         
  </tr>                         
  <tr>                         
    <td width="38%" height="31" bgcolor="#C0C0C0">班级<input type="text" name="bj1" size="20" value="<%=bj1%>"></td>                         
  </tr>                         
  <tr>                         
    <td width="38%" height="31" bgcolor="#C0C0C0">                         
      <p align="center"><font color="#FF33CC">全部清除</font></td>                         
    <td width="38%" height="31" bgcolor="#C0C0C0">                         
      <p align="center"><input type="radio" value="yes"  name="alldel">是&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                          
      <input type="radio" name="alldel" value="no" checked>否</td>                         
  </tr>                         
  <tr>                         
    <td width="24%" height="14" align="center" bgcolor="#666699"><font size="4" face="华文新魏">&nbsp;&nbsp;                          
      <input type="checkbox" name="tkgl3" value="ON">删除试卷库的试卷</font></td>                         
    <td width="38%" height="14" bgcolor="#C0C0C0">                         
      <p align="center">试卷名<input type="text" name="sjname" size="20" value="<%=sjname%>"></p>                         
    </td>                         
    <td width="38%" height="14" bgcolor="#C0C0C0" nowrap>                         
      <p align="center"><input type="radio" value="yes" name="clea" checked>是<input type="radio" name="clea" value="no">否                          
      <input type="submit" value="浏览" name="B5"></p>                         
      <p>　</td>                         
  </tr>                         
  </table>                         
  </center>                         
</div>                         
 
  <center> <input type="submit" value="执&nbsp;行&nbsp;命&nbsp;令" name="B5">  </center>    
 </form>                      
<p align="center"><font size="5" face="华文隶书" color="#000066"><span style="background-color: #C0C0C0"><a href="index.htm">返回到首页</a></span></font></p>                        
</body>                        
                        
</html>                        
