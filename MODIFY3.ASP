<!--#INCLUDE FILE="conn.asp" -->
<script language="javascript">
function check(myform)
{if(myform.question.value=="")
{alert("请输入题目");
return false;
}}
</script>
 <html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>asdfasdfasdfasdfsdfasdf</title>
</head>
<body>
<%dim exec,dh,tx,ts,na,m,sm, kssj,xm,xh,bj
   
function zh(fString)

   ' fString = replace(fString , "&gt;", ">")
    'fString = replace(fString, "&lt;", "<")
    'fString = replace(fString, "&nbsp;", " ")
    
    fString = Replace(fString, "&lt;BR&gt;", "<BR>")
    
    zh = fString
end function
 
  
   
   xh=request.querystring("xh")
  session("xh")=xh
  exec=session("exec") 
   sm=session("sm")
   dh=session("dh")
   kssj=session("kssj")
   kshm=session("kshm")
'set conn=server.createobject("adodb.connection")  
'conn.open "dsn=xiong1"  
set rs=server.createobject("adodb.recordset") 
set rs1=server.createobject("adodb.recordset") 
rs.open "select * from xscj where(kshm='"&kshm&"' and xh='"&xh&"')",conn,3,3 
xm=rs("xm")
bj=rs("bj")
rs.close
%> 
<form method="post" action="bcfs.asp"  name="myform"> 
<center><span style="background-color: #FFFF00"><font face="楷体_GB2312" color="#0000FF" size="5"><%=sm%></font></span> 
<p>
姓名:<%=xm%> &nbsp;&nbsp;                                                        
学号:<%=xh%>  &nbsp;&nbsp;                                                         
班级:<%=bj%>&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;                                                         
考试时间:<%=kssj%>分钟                                                           
</p> </center>                                                        
<table align="lift" border="1" width="100%" cellspacing="0" cellpadding="1">                                                            
<%                                            
                                            
rs.open exec,conn,3,3                                             
                                            
rs1.open "select * from da where(kshm='"&kshm&"' and xh='"&xh&"')",conn,3,3                                             
                                            
                                            
tx=""                                                            
 ts=0                                                            
%>                                                            
  <%do while not rs.eof  and not rs1.eof                                                                      
ts=ts+1                                                            
if tx<>rs("section") then                                                            
 ts=1                                                            
 tx=rs("section")                                                            
 %>                                                            
 <tr align="center"><td  width="50%" >                                                            
<span style="background-color: #FF3300"><font color="#000099"><font size="5" face="方正舒体">                                      
    <%=tx%></font></span>                                  
    　</td>                                        
 <td>                                  
 <%  m="本大题分数：*<input type='text' name='"+tx+"' size='4' value=0>"%>                                                           
   <span style="background-color: #66FF66"><font color="#6600FF" size="5" face="隶书">                                 
                                
  <%=m%>    </font></span>                                    
  　</td>                                                        
 </tr>                                                            
 <%end if%>                                                            
 <tr width="100%" ><td colspan="2"><%=ts&"."%><%=rs("que")%>　</td>                                                            
 </tr>                                                            
 <%if tx="选择题" then%>                                                            
 <tr align="center"><td width="50%">A:<%=rs("choicea")%></td><td width="50%">B:<%=rs("choiceb")%></td></tr>                                                            
  <tr align="center"><td width="50%">C:<%=rs("choicec")%></td><td width="50%">D:<%=rs("choiced")%></td></tr>                                                            
  <%end if%>                                                            
  <tr><td colspan="2">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
  <tr><td width="20%" align="right" bgColor="#669999">正确答案</td><td width="80%" ><%=zh(rs("answer"))%>　</td></tr>                                                            
  <tr><td align="right" bgColor="#666699"><font color="#FFFF00" face="楷体_GB2312" size="4"> 解答是：</font></td><td><%=rs1("da")%>　</td></tr>                                        
  </table>
  </td></tr>
  <tr><td colspan="2" align="right">                                      
  <font size="3" face="华文琥珀"><i><span style="background-color: #FFFF00">                              
  <input type="checkbox" name="xc" value="yes" <%if  ucase(rs1("da"))=ucase(rs("answer")) and rs1("da")<>"" then %> <% ="checked"%> <%end if%>>正确否？                                        
   </span></font>                                                          
<%  na=tx&ts %>                                                            
<%  m="本题分数：<input type='text' name='"+na+"' size='3' value=0>"%>                                                            
  <span style="background-color: #FF00FF"><font color="#FF0000" size="3" face="隶书">                            
  <%=m%> </font></span>                                                           
</td></tr>                                                       
                                                            
  <%rs.movenext%><%rs1.movenext%><%loop%>                                                   
<% rs.close                                                  
 rs1.close                                     
 conn.close                                                  
 %>　                                                                       
</table>      
  <p><center><input type="submit" name="submit" value="保存分数"></center></p>      
 </form>                                                           
  <center>         
                                                           
                                                            
<p> <a href=kh.asp>放弃并返回</a>                                                       
<a href=index.htm>回到首页</a>     </p>  </center>                                                
</body>                                                
</html>