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
<title>�½���ҳ 1</title>
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
    <b><font size="4" color="#FF00FF"><span style="background-color: #0000FF">��&nbsp;
    ��&nbsp;��&nbsp;¼</span></font></b></td>
  </tr>
  <tr>
    <td width="12%" height="16" align="center">׼��֤��</td>
    <td width="10%" height="16" align="center">����</td>
    <td width="10%" height="16" align="center">������</td>
    <td width="10%" height="16" align="center">�༶</td>
    <td width="10%" height="16" align="center">רҵ</td>
    <td width="5%" height="16" align="center">����</td>
    <td width="5%" height="16" align="center">��ѧ</td>
    <td width="5%" height="16" align="center">Ӣ��</td>
    <td width="5%" height="16" align="center">�ۺ�</td>
    <td width="10%" height="16" align="center">��</td>
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
     if zd="05��һ" then m1="<option selected>05��һ</option>" else m1=" <option>05��һ</option>"
     if zd="05�ƶ�" then m2="<option  selected>05�ƶ�</option>" else m2="<option >05�ƶ�</option>"
      if zd="05����" then m3="<option selected>05����</option>" else m3="<option>05����</option>"
      if zd="05����" then m4="<option selected>05����</option>" else m4="<option>05����</option>"
      if zd="04����" then m5="<option selected>04����</option>" else m5="<option>04����</option>"
      if zd="04��е" then m6="<option selected>04��е</option>" else m6="<option>04��е</option>"
      if zd="04��װ" then m7="<option selected>04��װ</option>" else m7="<option>04��װ</option></select>"
    m= "<select size='1' onChange='js("+chr(34)+na1+chr(34)+")' name='"+na+"' style='font-family: ����; font-size: 12pt' >"
  else
  if k=5 then
     m= "<select size='1' onChange='js("+chr(34)+na1+chr(34)+")' name='"+na+"' style='font-family: ����; font-size: 12pt' >"
     zd=trim(rs(k-1))
     if zd="�����" then m1="<option selected>�����</option>" else m1=" <option>�����</option>"
     if zd="����" then m2="<option  selected>����</option>" else m2="<option >����</option>"
     if zd="��е" then m3="<option selected>��е</option>" else m3="<option>��е</option>"
     if zd="��װ" then m4="<option selected>��װ</option>" else m4="<option>��װ</option></select>"
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
    <input type="submit"  value="�����¼" name="B1"></td>
    <td  height="16" align="center" colspan="5" ><a href="cjdisp1.htm">���ز�ѯҳ��</a> </td>
    
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
  <a href="modigkcj.asp?jlh=-1">��������</a><span lang="en-us">&nbsp; </span><a href="modigkcj.asp?jlh=30">����30��</a>
 <span lang="en-us">&nbsp; </span><a href="modigkcj.asp?jlh=50">����50��</a><span lang="en-us">&nbsp;
 </span><a href="modigkcj.asp?jlh=100">����100��</a> <span lang="en-us">&nbsp;</span><a href="modigkcj.asp?jlh=200">����200��</a>
 <span lang="en-us">&nbsp;</span>
 <br><a href="modigkcj.asp?jlh=0">����ʼ</a><span lang="en-us">&nbsp;&nbsp;&nbsp;
 </span><a href="modigkcj.asp?jlh=-2">������</a><span lang="en-us">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 </span><a href="modigkcj.asp?jlh=-50">����50�� </a> <span lang="en-us">&nbsp;</span><a href="modigkcj.asp?jlh=-3">��������</a>
 
</body>

</html>