<%
''''''''----====    lzhiy0816 本地站点     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 网上空间     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''
%>

<!--#include file="conn.asp"-->
<%    ''''''''====《功能：用户链接数据添加到数据库》(lzhiy0816 design)====''''''''

xingming=request.form("xingming")    ''''''''从表单获取数据'''''''
password=request.form("password")
mingchen=request.form("mingchen")
shuoming=request.form("shuoming")
lianjie=request.form("lianjie")
logoimg=request.form("logoimg")

exec="select xingming from logo"    ''''''''查寻xingming字段记录(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
xingming1=""                        ''''''''先将xingming置为空（lzhiy0816 design)''''''''
do while not rs.eof                 ''''''''如果提交的用户名在记录里存在，则将xingming1置为记录值（lzhiy0816 design)''''''''
  if xingming=rs("xingming") then
  xingming1=rs("xingming")
  end if
rs.movenext
loop
if xingming1<>"" then
response.redirect "logoback.asp?err="&xingming1&""
rs.close
set rs=nothing
else

''''''''数据添加到数据库'''''''
exec="insert into logo(xingming,password,mingchen,shuoming,lianjie,logoimg) values('"+xingming+"','"+password+"','"+mingchen+"','"+shuoming+"','"+lianjie+"','"+logoimg+"')"
conn.execute exec
conn.close
set conn=nothing
response.redirect "logo.asp"
end if
%>
