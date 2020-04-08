<%
''''''''====《管理员密码检查logo》(lzhiy0816 design)====''''''''
dim inputxm,inputpw    ''''''''获取lzhiy0816数据(lzhiy0816 design)''''''''
inputxm=request.form("xingming")
inputpw=request.form("password")
id=request.querystring("id")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select * from admin where id=1"    ''''''''查寻lzhiy0816的id号(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
''''''''如果是lzhiy0816则进入《网 站 管 理 员 管 理 窗 口》界面(lzhiy0816 design)''''''''
if inputxm=rs("xingming")  and inputpw=rs("password") and rs("password")<>"" then
session("lzychecked")="yes"
session("lzycheck")="right"
response.Redirect "logolzymdflogin.asp?id="&id&""
''''''''如果不是lzhiy0816回到《网 站 管 理 员 入 口》界面(lzhiy0816 design)''''''''
else
session("lzychecked")="no"
session("lzycheck")="wrong"
response.Redirect "logolzymdf.asp?id="&id&""
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
