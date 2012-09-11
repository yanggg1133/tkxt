<!--#INCLUDE FILE="conn.asp" -->
<%set rs=server.createobject("adodb.recordset")
rs.open "select * from gkcj",conn,3,3
%>
<% dim jl,jlh,m1,m2,m3,m4,m5,m6,m7,row,k,na,na1,m,dq
if session("dq")="" then
 dq=0
 else
 dq=session("dq")
end if
jlh=request.querystring("jlh")
jl=cINT(jlh)
if jl=0 then dq=0
if jl=-2 then
jl=0
dq=rs.recordcount-8
if dq<0 then  dq=0
end if
if jl=-3 then 
 dq=dq-8
 if dq<0 then dq=0
end if 
if jl=-1 then jl=0
jl=jl+dq+1
if jl<1 then jl=1
rs.move jl

%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新建网页 1</title>
<SCRIPT language=JavaScript>
function js(x)
{ 

 xx="document.form1."+x
 eval(xx).checked=1
 
}
</script>
</head>

<body>
 <form  name="form1" method="POST" action="modigkcj1.asp">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="488" id="AutoNumber1" height="89">
  <tr>
    <td width="489" align="center" colspan="10" height="16">
    <b><font size="4" color="#FF00FF"><span style="background-color: #0000FF">修&nbsp;
    改&nbsp;记&nbsp;录</span></font></b></td>
  </tr>
  <tr>
    <td width="12%" height="16" align="center">准考证号</td>
    <td width="10%" height="16" align="center">姓名</td>
    <td width="10%" height="16" align="center">报名号</td>
    <td width="10%" height="16" align="center">班级</td>
    <td width="10%" height="16" align="center">专业</td>
    <td width="5%" height="16" align="center">语文</td>
    <td width="5%" height="16" align="center">数学</td>
    <td width="5%" height="16" align="center">英语</td>
    <td width="5%" height="16" align="center">综合</td>
    <td width="10%" height="16" align="center">？</td>
  </tr>
  <%
 row=1
%>
  <%do while not rs.eof  and row<=8  
  %>
  <tr>
<% k=1
    do while k<=10
     na="tt"&row&k
     na1="tt"&row&"10"
   if k=4 then
     zd=trim(rs(k-1))
     if zd="05计一" then m1="<option selected>05计一</option>" else m1=" <option>05计一</option>"
     if zd="05计二" then m2="<option  selected>05计二</option>" else m2="<option >05计二</option>"
      if zd="05计三" then m3="<option selected>05计三</option>" else m3="<option>05计三</option>"
      if zd="05计四" then m4="<option selected>05计四</option>" else m4="<option>05计四</option>"
      if zd="04旅游" then m5="<option selected>04旅游</option>" else m5="<option>04旅游</option>"
      if zd="04机械" then m6="<option selected>04机械</option>" else m6="<option>04机械</option>"
      if zd="04服装" then m7="<option selected>04服装</option>" else m7="<option>04服装</option></select>"
    m= "<select size='1' onChange='js("+chr(34)+na1+chr(34)+")' name='"+na+"' style='font-family: 宋体; font-size: 12pt' >"
  else
  if k=5 then
     m= "<select size='1' onChange='js("+chr(34)+na1+chr(34)+")' name='"+na+"' style='font-family: 宋体; font-size: 12pt' >"
     zd=trim(rs(k-1))
     if zd="计算机" then m1="<option selected>计算机</option>" else m1=" <option>计算机</option>"
     if zd="旅游" then m2="<option  selected>旅游</option>" else m2="<option >旅游</option>"
     if zd="机械" then m3="<option selected>机械</option>" else m3="<option>机械</option>"
     if zd="服装" then m4="<option selected>服装</option>" else m4="<option>服装</option></select>"
   else
    if k=10 then
    m="<input type='checkbox' size='8' name='"+na+"' value='ON'>"
    else
   if k>=6 and k<=9 then
   m="<input type='text'  onChange='js("+chr(34)+na1+chr(34)+")' name='"&na&"' size='4' value="&rs(k-1)&">"
   else
   m="<input type='text'  onChange='js("+chr(34)+na1+chr(34)+")' name='"&na&"' size='10' value="&rs(k-1)&">"
    end if  
  end if
     end if
     end if
     
  %>
  
    <td  height="21">
 
     <%=m%>
     <% if k=4 then %>
      <%=m1%>
     <%=m2%>
     <%=m3%>
     <%=m4%>
     <%=m5%>
     <%=m6%>
     <%=m7%>
     <% end if %>
     
     <% if k=5 then %>
      <%=m1%>
     <%=m2%>
     <%=m3%>
     <%=m4%>
     <% end if %>
     

        </td>
        
        <%
        k=k+1
        loop%>
        </tr>
       <%rs.movenext%><%
       row=row+1
       loop%> 
    <%session("dq")=jl+row-2
        session("dq1")=jl
          session("gs")=row-1
 %>
  <tr>
    <td  height="16" align="center" colspan="5" >
    <input type="submit"  value="保存记录" name="B1"></td>
    <td  height="16" align="center" colspan="5" ><a href="cjdisp1.htm">返回查询页面</a> </td>
    
    </tr>
  <tr>
    <td  height="16" align="center" >start:</td>
    
     
    <td  height="16" align="left" > <%=jl %></td>
    <td  height="16" align="center" >end:</td>
    <td  height="16" align="left" ><%=session("dq") %> </td>
    <td  height="16" colspan="2" align="center" >record:</td>
    <td  height="16" align="left" colspan="2" > <%=session("dq") %>sum:</td>

    <td  height="16" align="left" colspan="2" ><%=rs.recordcount%> </td>
    </tr>
    </table>
</form>
  <a href="modigkcj.asp?jlh=-1">继续下移</a><span lang="en-us">&nbsp; </span><a href="modigkcj.asp?jlh=30">下移30条</a>
 <span lang="en-us">&nbsp; </span><a href="modigkcj.asp?jlh=50">下移50条</a><span lang="en-us">&nbsp;
 </span><a href="modigkcj.asp?jlh=100">下移100条</a> <span lang="en-us">&nbsp;</span><a href="modigkcj.asp?jlh=200">下移200条</a>
 <span lang="en-us">&nbsp;</span>
 <br><a href="modigkcj.asp?jlh=0">到开始</a><span lang="en-us">&nbsp;&nbsp;&nbsp;
 </span><a href="modigkcj.asp?jlh=-2">到结束</a><span lang="en-us">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 </span><a href="modigkcj.asp?jlh=-50">上移50条 </a> <span lang="en-us">&nbsp;</span><a href="modigkcj.asp?jlh=-3">继续上移</a>
 
</body>

</html>