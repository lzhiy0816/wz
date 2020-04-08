<!--#include file="conn.asp"-->
<%
''''''''====《将id号用户提交修改的数据，改写进数据库》(lzhiy0816 design)====''''''''

exec="select * from admin where id=1"    ''''''''查寻lzhiy0816的数据(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
xingming=rs("xingming")
password=rs("password")
rs.close
set rs=nothing

exec="select * from userleave where id="&request.form("id")    ''''''''查寻id号提交用户的数据(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
''''''''如果是lzhiy0816则显示lzhiy0816头像，否则显示用户提交的头像(lzhiy0816 design)''''''''
if rs("xingming")=xingming and rs("password")=password then
rs("face")="99"
else rs("face")=request.form("face")
end if
rs("dizhi")=request.form("dizhi")
rs("email")=request.form("email")
rs("zhuye")=request.form("zhuye")
rs("icon")=request.form("icon")
liuyan=request.form("liuyanneirong")
liuyanneirong=Replace(liuyan,vbcrlf,"<Br>")
liuyanneirong=Replace(liuyanneirong," ","&nbsp;")
rs("liuyanneirong")=liuyanneirong
rs.update
rs.close
set rs=nothing

conn.close
set conn=nothing
response.redirect "leaveword.asp"
%>
