<!--#include file="conn.asp"-->
<%
''''''''====����id���û��ύ�޸ĵ����ݣ���д�����ݿ⡷(lzhiy0816 design)====''''''''

exec="select * from admin where id=1"    ''''''''��Ѱlzhiy0816������(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
xingming=rs("xingming")
password=rs("password")
rs.close
set rs=nothing

exec="select * from userleave where id="&request.form("id")    ''''''''��Ѱid���ύ�û�������(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
''''''''�����lzhiy0816����ʾlzhiy0816ͷ�񣬷�����ʾ�û��ύ��ͷ��(lzhiy0816 design)''''''''
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
