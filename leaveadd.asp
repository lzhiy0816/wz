<!--#include file="conn.asp"-->
<%    ''''''''====�����ܣ�����������ӵ����ݿ⡷(lzhiy0816 design)====''''''''
xingming=request.form("xingming")    ''''''''�ӱ���ȡ����'''''''
password=request.form("password")
password1=request.form("password1")

exec="select * from admin where id=1"    ''''''''��Ѱlzhiy0816��id��(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
''''''''�����lzhiy0816����ʾlzhiy0816ͷ�񣬷�����ʾ�û��ύ��ͷ��(lzhiy0816 design)''''''''
if xingming=rs("xingming")  and password=rs("password") and rs("password")<>"" then
face="99"
else face=request.form("face")
end if
dizhi=request.form("dizhi")
email=request.form("email")
zhuye=request.form("zhuye")
icon=request.form("icon")
liuyan=request.form("liuyanneirong")
liuyanneirong=Replace(liuyan,vbcrlf,"<Br>")    ''''''''�س�����'''''''
liuyanneirong=Replace(liuyanneirong," ","&nbsp;")    ''''''''�ո�'''''''
userip=Request.ServerVariables("Remote_Addr")
if password<>password1 then
response.Redirect "leave.htm"
else
''''''''������ӵ����ݿ�'''''''
exec="insert into userleave(xingming,password,face,dizhi,email,zhuye,icon,liuyanneirong,userip) values('"+xingming+"','"+password+"',"+face+",'"+dizhi+"','"+email+"','"+zhuye+"',"+icon+",'"+liuyanneirong+"','"+userip+"')"
conn.execute exec
rs.Close
set rs=nothing
conn.close
set conn=nothing
response.redirect "leaveword.asp"
end if
%>
