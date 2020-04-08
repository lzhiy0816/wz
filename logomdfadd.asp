<!--#include file="conn.asp"-->
<%
''''''''====《将用户提交修改的链接数据，改写进数据库》(lzhiy0816 design)====''''''''

id=request.querystring("id")
password=request.form("password")

if "提交"=Request.Form("Submit") then    ''''''''如果提交值为"提交"(lzhiy0816 design)''''''''
exec="select * from logo where id="&id    ''''''''查寻id号用户提交的链接数据(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3

''''''''修改链接数据，改写进数据库(lzhiy0816 design)''''''''
rs("xingming")=request.form("xingming")
if password<>"" then
rs("password")=request.form("password")
end if
rs("mingchen")=request.form("mingchen")
rs("shuoming")=request.form("shuoming")
rs("lianjie")=request.form("lianjie")
rs("logoimg")=request.form("logoimg")

rs.update
rs.close
set rs=nothing
end if

if "删除"=Request.Form("Submit") then    ''''''''如果提交值为删除(lzhiy0816 design)''''''''
exec="delete * from logo where id="&id    ''''''''lzhiy0816删除某id号的留言内容(lzhiy0816 design)''''''''
conn.execute exec
end if

conn.close
set conn=nothing
response.redirect "logo.asp"
%>
