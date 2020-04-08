<%
''''''''----====    lzhiy0816 本地站点     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 网上空间     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''
%>

<!--#include file="conn.asp"-->
<%    ''''''''====《功能：留言数据添加到数据库》(lzhiy0816 design)====''''''''
xingming=request.form("xingming")    ''''''''从表单获取数据'''''''
password=request.form("password")
password1=request.form("password1")

exec="select * from admin where id=1"    ''''''''查寻lzhiy0816的id号(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
adminface=rs("face")
''''''''如果是lzhiy0816则显示lzhiy0816头像，否则显示用户提交的头像(lzhiy0816 design)''''''''
if xingming=rs("xingming")  and password=rs("password") and rs("password")<>"" then
face=adminface
else face=request.form("face")
end if
dizhi=request.form("dizhi")
email=request.form("email")
zhuye=request.form("zhuye")
icon=request.form("icon")
liuyan=request.form("liuyanneirong")
liuyanneirong=Replace(liuyan,vbcrlf,"<Br>")    ''''''''回车换行'''''''
liuyanneirong=Replace(liuyanneirong," ","&nbsp;")    ''''''''空格'''''''
userip=Request.ServerVariables("Remote_Addr")
if password<>password1 then
response.Redirect "leave.htm"
else
''''''''数据添加到数据库'''''''
exec="insert into userleave(xingming,password,face,dizhi,email,zhuye,icon,liuyanneirong,userip) values('"+xingming+"','"+password+"',"+face+",'"+dizhi+"','"+email+"','"+zhuye+"',"+icon+",'"+liuyanneirong+"','"+userip+"')"
conn.execute exec
rs.Close
set rs=nothing
conn.close
set conn=nothing
response.redirect "leaveword.asp"
end if
%>
