<!--#include file="conn.asp"-->
<%
''''''''====������ĳ�û��ĵ��������1�󱣴棬Ȼ���lzhiy0816����ҳ��(lzhiy0816 design)====''''''''

xingming=request.querystring("user")

exec="select * from logo where xingming='"&xingming&"'"    ''''''''��Ѱ������û�����������(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3

''''''''�޸��������ݣ���д�����ݿ�(lzhiy0816 design)''''''''
application.lock
rs("countme")=rs("countme")+1
rs.update
application.unlock

response.redirect "index.htm"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
