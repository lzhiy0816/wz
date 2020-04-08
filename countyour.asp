<!--#include file="conn.asp"-->
<%
''''''''====《查找某用户的点击数，加1后保存，然后打开该用户的主页》(lzhiy0816 design)====''''''''

xingming=request.querystring("user")

exec="select * from logo where xingming='"&xingming&"'"    ''''''''查寻被点击用户的链接数据(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3

''''''''修改链接数据，改写进数据库(lzhiy0816 design)''''''''
application.lock
rs("countyour")=rs("countyour")+1
rs.update
application.unlock

response.redirect rs("lianjie")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
