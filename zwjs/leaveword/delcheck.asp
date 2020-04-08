<%
''''''''====《管理员密码检查》(lzhiy0816 design)====''''''''
dim inputxm2,inputpw2    ''''''''获取lzhiy0816数据(lzhiy0816 design)''''''''
inputxm2=request.form("xingming2")
inputpw2=request.form("password2")
id=request.querystring("id")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select * from admin where id=1"    ''''''''查寻lzhiy0816的id号(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
''''''''如果是lzhiy0816则进入《网 站 管 理 员 管 理 窗 口》界面(lzhiy0816 design)''''''''
if inputxm2=rs("xingming")  and inputpw2=rs("password") and rs("password")<>"" then
session("checked")="yes2"
session("check")="right2"
response.Redirect "delorans.asp?id="&id&""
''''''''如果不是lzhiy0816回到《网 站 管 理 员 入 口》界面(lzhiy0816 design)''''''''
else
session("checked")="no2"
session("check")="wrong2"
response.Redirect "dellogin.asp?id="&id&""
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
