<%@language=vbscript%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 3</title>
</head>

<body>
<%dim sm,kshm,bj
sm=request.querystring("sm")
kshm=request.querystring("kshm")
bj=request.querystring("bj")
if sm="" then
sm=session("sm")
end if
if kshm="" then
kshm=session("kshm")
end if
if bj="" then
bj=session("bj")
end if
session("sm")=sm
session("bj")=bj
session("kshm")=kshm
%>
<p align="center"><font face="楷体_GB2312" size="6" color="#FFFF00"><span style="background-color: #FF0066">欢 
迎 进 入 查 询 系 统</span></font></p>
<div align="center">
  <center><form method="POST" name="form1" action=cxjg.asp>
  
  <table border="1" width="100%" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#66FF66" >
    <tr>
      <td width="100%" colspan="3" height="24">
        <p align="center"><b><font size="5" face="仿宋_GB2312">选择科目&nbsp;&nbsp;&nbsp;  
        </font><a href="xckmll.asp"><font size="3" face="楷体_GB2312">浏览</font></a><font size="3" face="楷体_GB2312">选择</font></b></td>                         
    </tr>                         
    <td colspan="3"  align="center"> 
<span style="background-color: #FFFF00"><font face="楷体_GB2312" color="#0000FF" size="3"><%=sm%>  
</font></span> 
 
</td></tr>   
  <tr>                         
      <td width="31%" rowspan="3" height="17">                         
        <p align="center"><font size="5" face="华文隶书" color="#000066"><input type="checkbox" name="gr" value="ON" checked>个人查询</font></td>                         
      <td width="69%" height="31"  align="center" colspan="2">                         
                                 
         姓&nbsp; 名：<input type="text" name="name" size="10" maxlength="12">                        
                                
      </td>                        
    </tr>                        
    <tr>                        
      <td width="69%" height="31" align="center" colspan="2">学&nbsp; 号：<input type="text" name="xh" size="10" maxlength="12"></td>                        
    </tr>                        
    <tr>                        
      <td width="69%" height="31" align="center" colspan="2">                      
        <p align="center">所有科目:<input type="checkbox" name="all" value="ON"></td>                        
    </tr>                        
    <tr>                        
      <td width="31%" height="64" rowspan="4">                     
        <p align="center"><font size="5" face="华文隶书" color="#000066"><input type="checkbox" name="bj" value="ON" checked>班级查询</font></td>                      
      <td width="69%" height="31"  align="center" colspan="2">班级名：<input type="text" name="bjname"  value='<%=bj%>'size="10" maxlength="18"></td>                      
    </tr>                      
    <tr>                      
      <td width="69%" height="31" colspan="2">                   
        <p align="center"><font size="3" face="楷体_GB2312" color="#0000FF"><b><a href="bjll.asp">浏览选择班级</a></b></font></td>                     
    </tr>                     
    <tr>                     
      <td width="12%" height="62" rowspan="2">                  
        <p align="center"><b>分数段</b></td>                     
      <td width="57%" height="31">30以下：<input type="checkbox" name="cc1" value="on" checked>&nbsp;                   
        30－59：<input type="checkbox" name="cc2" value="on" checked>&nbsp; 60－79：<input type="checkbox" name="cc3" value="on" checked>&nbsp;80－100：<input type="checkbox" name="cc4" value="on" checked></td>                     
    </tr>                     
    <tr>                     
      <td width="57%" height="31" align="center"><font face="仿宋_GB2312"><u>指定分数段</u></font>：                  
      <input type="text" name="start" size="4" maxlength="4" value=0> －<font color="#FF00FF">to</font>－                 
      <input type="text" name="end" size="4" maxlength="4" value=-1></td>                     
    </tr>                     
  </table>        
          
  <p>  <input type="submit" name="submit" value="查看查询结果">  </p>       
        
        
  </form>                     
  </center>                     
</div>                     
                     
<p align="center"><a href="kh.asp">返回到考试页面</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="index.htm">返回首页</a></p>       
                   
</body>                   
                   
</html>                   
















