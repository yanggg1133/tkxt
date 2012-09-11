<%@language=vbscript %>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 4</title>
</head>

<body>
<%dim x(2,2)
x(0,0)=1
x(0,1)="b"
response.write x(0,0)
response.write x(0,1)
%>
<center>
<p>本套试卷不能被保存...请与发卷人联系....
 <p><a href=kh.asp>回到首页</a></p></center>
 <%else
  tx(0,1)=rs("xc")
  tx(1,1)=rs("wxtk")
  tx(2,1)=rs("tk")
  tx(3,1)=rs("pd")
  tx(4,1)=rs("jd")
  tx(5,1)=rs("js")
  tx(6,1)=rs("cx")
  tx(7,1)=rs("cxjg")
  end if
  rs.close%
  response.write tx(0,1)
  %>
<p> 
    <center>
    <p>用户名为空，不能被追加</center>
    <%end if%> 
 <%else%> 
 <center>不能追加用户，该用户已存在</center> 
 <%end if%> 
 <%rs.close 
   end if%> 
    
    
<%if user1="3" then 
  dim  userdel,sfdel 
  userdel=request.form("userdel") 
  sfdel=request.form("sfdel") '删除用户浏览 
     if sfdel="yes" then 
     rs.open exec,conn,3,3%> 
     <table align="center" width="80%"> 
    <tr><td colspan="2">删除用户浏览结果</td></tr> 
        <%if  not rs.eof then%> 
          <tr><td width="50%">用户名</td><td>用户密码</td></tr> 
              <%do while not rs.eof%> 
                <tr><td><a href=xtgl.asp?username=<%=rs(2)%>><%=rs(2)%></a></td> <td><%=rs(3)%></td></tr> 
              <%rs.movenext%><%loop%> 
      </table> 
        <%else %> 
       <center><br><p>该数据库无用户</p></br></center>      
成功删除      
       <%end if%> 
     <%rs.close%> 
     <%else%> 
     <%conn.execute "delete from tblusr  where usrid='"+userdel+"'"   %> 
     <center>成功删除</center> 
     <%end if%> 
<%end if%> 
 
 
 
<%if user1="4" then %> 
<%  dim  username1,userpass1 
  username1=request.form("username") 
  userpass1=request.form("userpass") 
  modill=request.form("modill") '修改用户浏览 
     if modill="yes" then 
     rs.open exec,conn,3,3%> 
     <table align="center" width="80%"> 
    <tr><td colspan="2">修改用户密码浏览结果</td></tr> 
        <%if  not rs.eof then%> 
          <tr><td width="50%">用户名</td><td>用户密码</td></tr> 
              <%do while not rs.eof%> 
                <tr><td><a href=xtgl.asp?username=<%=rs(2)%>&userpass=<%rs(3)%>><%=rs(2)%></a></td> <td><%=rs(3)%></td></tr> 
              <%rs.movenext%><%loop%> 
      </table> 
        <%else %> 
       <center><br><p>该数据库无用户</p></br></center>      
成功地修改密码用户不存在，不能修改密码      
       <%end if%> 
     <%rs.close%> 
     <%else%> 
     <% m="select * from tblusr where(usrid='"+username1+"')" 
      rs.open m,conn,3,3 
      if not rs.eof then 
       rs("usrpwd")=userpass1 
       rs.update%> 
        <center>成功地修改密码</center> 
      <%else%> 
        <center>用户不存在，不能修改密码</center> 
      <%end if 
       rs.close 
      end if%> 
<%end if%> 
 
</body> 
 
</html> 
<%rs.close 
tx(0,0)="选择题" 
tx(1,0)="完型填空" 
tx(2,0)="填空题" 
tx(3,0)="判断题" 
tx(4,0)="简答题" 
tx(5,0)="计算题" 
tx(6,0)="程序题" 
tx(7,0)="程序结果" 
m="select * from sj where(kshm='"+kshm+"' and dh='"+dh+"')" 
rs.open m,conn,3,3 
if rs.eof then 
%>   
 
 
 
 
<%if user1="2" then 
 dim useradd,pass 
 useradd=request.form("useradd") 
 pass=request.form("pass") 
 m="select * form tblusr where(usrid='"+useradd+"' and usrpwd='"+pass+"')" 
 rs.open m,conn,3,3 
 if rs.eof then 
    if useradd<>"" then 
      rs.addnew 
      rs("usrid")=useradd 
      rs("usrpwd")=pass 
      rs.update 
    else%> 





 <%else%>
    <% dim sjname
     sjname=request.form("sjname")
     if  sjname<>"" then
    m="delete from sj where sm='"+sm+"'"
      else 
    if sjname="*" then
    m="delete from sj"  
    end if
   end if
    if sjname<>"" then
    conn.execute m%>
    <center>成功删除了规定的学生答案</center>
    <%else%>
    <center>未删除,可能是没有输入试卷名或输错了试卷名</center>
    <%end if%>

