<%
''''''''====《用户密码的验证》(lzhiy0816 design)====''''''''
dim inputxm,inputpw
inputpw=request.form("password")
id=request.querystring("id")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select * from userleave where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn    ''''''''从数据库获取id号用户的数据''''''''
if inputpw=rs("password") and rs("password")<>"" then
session("checked")="yes"
session("check")="right"
response.Redirect "modify.asp?id="&id&""    ''''''''密码正确，则进入《请 修 改 您 的 留 言 及 相 关 信 息》界面(lzhiy0816 design)''''''''
else
session("checked")="no"
session("check")="wrong"
response.Redirect "mdflogin.asp?id="&id&""    ''''''''密码不对，则返回《用 户 修 改 留 言 入 口》界而(lzhiy0816 design)'''''''
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
