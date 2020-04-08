<%
''''''''====《网站站长密码检查》(lzhiy0816 design)====''''''''
dim inputxm,inputpw    ''''''''获取网站站长输入的网站让长数据(lzhiy0816 design)''''''''
inputxm=request.form("xingming")
inputpw=request.form("password")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%

exec="select * from admin where id=1"    ''''''''获取lzhiy0816的"xingming"和"password"值(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
adminxm=rs("xingming")
adminpw=rs("password")
rs.Close
set rs=nothing

''''''''如果是lzhiy0816，则回到logo.asp的lzhhiy0816管理界面(lzhiy0816 design)''''''''
if inputxm=adminxm and inputpw=adminpw then
response.Redirect "logolzy.asp"
else
exec="select * from logo where xingming='"&inputxm&"'and password='"&inputpw&"'"     ''''''''查寻网站站长的用户名和密码(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
''''''''如果是网站站长，则进入《修改链接管理窗口》界面(lzhiy0816 design)''''''''
do while not rs.eof 
inputxm1=rs("xingming")
inputpw1=rs("password")
rs.movenext 
loop 
if inputxm1<>"" then
session("checked")="yes2"
session("check")="right2"
response.Redirect "logomdflogin.asp?user="&inputxm1
''''''''如果不是网站站长,回到《网 站 站 长 验 证》界面(lzhiy0816 design)''''''''
else
session("checked")="no2"
session("check")="wrong2"
response.Redirect "logomdf.asp"
end if
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%>
 
