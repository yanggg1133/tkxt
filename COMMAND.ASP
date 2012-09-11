<!--#INCLUDE FILE="conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>考试号码的浏览</title>
</head>

<body>

<p align="center"><font size="5" face="隶书">系统管理处理结果</font></p>

<% dim userlist,modill,user1,user,m
user=request.form("user")  '用户管理大项
tkgl=request.form("tkgl") '题库管理大项
user1=request.form("user1")
'set conn=server.createobject("adodb.connection")                     
'conn.open "dsn=xiong1"                     
set rs=server.createobject("adodb.recordset") 

if user="ON" then 
exec= "select * from tblusr"
if user1="1" then 
 userlist= request.form("userlist") '用户浏览的通配符
 m="select * from tblusr where(usrid like '%"&userlist&"%')"
 rs.open m,conn,3,3
 if  not rs.eof then%>
<table align="center" width="80%">
<tr><td colspan="2" align="center">用户浏览结果</td></tr>
<tr><td width="50%">用户名</td><td>用户密码</td></tr>
<%do while not rs.eof%>
<tr><td><%=rs(1)%></td> <td><%=rs(2)%></td></tr>
<%rs.movenext%><%loop%>
</table>
<%else %>
<center><br><p>该数据库无用户</p></br></center>
<%end if%>

<%rs.close
                    
end if%>


<%if user1="2" then 
 dim useradd,pass 
 useradd=request.form("useradd") 
 pass=request.form("pass") 
 m="select * from tblusr where(usrid='"+useradd+"')" 
 rs.open m,conn,3,3 
 if rs.eof then 
    if useradd<>"" then 
      rs.addnew 
      rs("usrid")=useradd 
      rs("usrpwd")=pass 
      rs.update 
    else%>
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
                <tr><td><a href=xtgl.asp?username=<%=rs(1)%>><%=rs(1)%></a></td> <td><%=rs(2)%></td></tr> 
              <%rs.movenext%><%loop%> 
      </table> 
        <%else %> 
        </table>
       <center><br><p>该数据库无用户</p></center>      
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
                <tr><td><a href=xtgl.asp?username=<%=rs(1)%>&userpass=<%=rs(2)%>><%=rs(1)%></a></td> <td><%=rs(2)%></td></tr> 
              <%rs.movenext%><%loop%> 
      </table> 
        <%else %>
        </table> 
       <center><br><p>该数据库无用户</p></center>      
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
<%end if %>



<%
'response.write request.form("tkgl1")
'response.end
%>
<%dim tkgl,tkgl1,tkgl2,tkgl3
tkgl=request.form("tkgl")
tkgl1=request.form("tkgl1")
tkgl2=request.form("tkgl2")
tkgl3=request.form("tkgl3")
if tkgl="ON" then
    if tkgl1="ON" then
       dim kmll
      
       kmll=request.form("kmll")
       if kmll="yes" then
          'm="select * from xscj,sjdh where(xscj.dh=sjdh.dh)"
           'm="select  distinct sjdh.sm,xscj.bj from xscj right join sjdh on xscj.dh=sjdh.dh order by sm"
           m="select * from sjdh,(select distinct xscj.kshm,xscj.dh,xscj.bj from xscj) where(xscj.dh=sjdh.dh)" 
           rs.open m,conn,3,3
           %>
            <table align="center" width="80%"> 
             <tr><td colspan="2" align="center">科目及班级浏览结果</td></tr> 
           <%if  not rs.eof then%> 
           <tr><td width="50%">科目名</td><td>班级名</td></tr> 
              <%do while not rs.eof%> 
                <tr><td><a href=xtgl.asp?km=<%=rs("sm")%>&bj=<%=rs("bj")%>><%=rs("sm")%>(<%=rs("kshm")%>)</a></td> <td><%=rs("bj")%></td></tr> 
              <%rs.movenext%><%loop%> 
          </table> 
          <%else %> 
           </table>
           <center><br><p>该数据库无用户</p></center>      
          <%end if%> 
     <%rs.close%>
    <%else%>
    <% dim bj,km
     bj=request.form("bj")
     km=request.form("km")
    if bj<>"" and km<>"" then
    m="delete from da where xh in (select xh from xscj where bj='"+bj+"')"+" and dh=(select dh from sjdh where sm='"+km+"')"
     else 
     if bj<>"" then 
      m="delete from da where xh in (select xh from xscj where bj='"+bj+"')"
     else 
       if km<>"" then
        m="delete from da where dh=(select dh from sjdh where sm='"+km+"')"
       else
        m="delete from da"
       end if
     end if
   end if
    
    conn.execute m
   %> 
    <center>成功删除了规定的学生答案</center>
    
    <%end if%>
     
 <% end if



 if tkgl2="ON" then
       dim xhll
       xhll=request.form("xhll")
       if xhll="yes" then
           m="select * from xscj"
           rs.open m,conn,3,3
           %>
            <table align="center" width="80%"> 
             <tr><td colspan="2" align="center">班级和学生学号浏览结果</td></tr> 
           <%if  not rs.eof then%> 
           <tr><td width="50%">班级名</td><td>学号</td></tr> 
              <%do while not rs.eof%> 
                <tr><td><a href=xtgl.asp?bj1=<%=rs("bj")%>&xh=<%=rs("xh")%>><%=rs("bj")%></a></td> <td><%=rs("xh")%></td></tr> 
              <%rs.movenext%><%loop%> 
          </table> 
          <%else %> 
           </table>
           <center><br><p>该数据库无学生成绩档案</p></center>      
          <%end if%> 
     <%rs.close%>
    <%else%>
    <% dim bj1,xh,alldel
     alldel=request.form("alldel")
     bj1=request.form("bj1")
     xh=request.form("xh")
    if bj1<>"" and xh<>"" then
    m="delete from xscj where bj='"+bj1+"' and xh='"+xh+"'"
    else 
     if bj1<>"" then 
      m="delete from xscj where bj='"+bj1+"'"
     else 
       if xh<>"" then
        m="delete from xscj where xh='"+xh+"'"
      end if
     end if
   end if %>
   <% if alldel="yes" then
     conn.execute "delete from xscj"
     %>
        <center>成功删除了全部学生档案</center>
    <%else
     if bj1<>"" or xh<>"" then
     conn.execute m
     %>
     <center>成功删除了规定的学生档案</center>
     <%else%>
     <center>未删除任何学生档案</center>
     <%end if
     end if %>    
  <%end if%>
  <% end if%>




<%    if tkgl3="ON" then
       dim clea
       clea=request.form("clea")
       if clea="yes" then
           m="select * from sj"
           rs.open m,conn,3,3
           %>
            <table align="center" width="80%"> 
             <tr><td colspan="2" align="center">试卷库的试卷浏览结果</td></tr> 
           <%if  not rs.eof then%> 
           <tr><td width="50%" colspan="2">试卷名</td></tr> 
              <%do while not rs.eof%> 
              <tr><td colspan="2"><a href=xtgl.asp?sm=<%=rs("sm")%>><%=rs("sm")%></a></td> </tr> 
              <%rs.movenext%><%loop%> 
          </table> 
          <%else %> 
           </table>
           <center><br><p>老师没有出任何试卷</p></center>      
          <%end if%> 
     <%rs.close%>
     <%else%>
    <% dim sjname
     sjname=request.form("sjname")
     if  sjname<>"" then
    m="delete from sj where sm='"+sjname+"'"
      end if 
    if sjname="*" then
    m="delete from sj"  
    
   end if
    if sjname<>"" then
    conn.execute m%>
    <center>成功删除了规定的试卷名</center>
    <%else%>
    <center>未删除,可能是没有输入试卷名或输错了试卷名</center>
    <%end if%>
    
    <%end if%>
     
 <% end if%>








<%end if
%>



<br><p><center><a href=xtgl.asp>直接返回</a></center></p>
</body>

</html>






















































































































































