<%
''''''''====���û��������֤��(lzhiy0816 design)====''''''''
dim inputxm,inputpw
inputpw=request.form("password")
id=request.querystring("id")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select * from userleave where id="&request.querystring("id")
set rs=server.createobject("adodb.recordset")
rs.open exec,conn    ''''''''�����ݿ��ȡid���û�������''''''''
if inputpw=rs("password") and rs("password")<>"" then
session("checked")="yes"
session("check")="right"
response.Redirect "modify.asp?id="&id&""    ''''''''������ȷ������롶�� �� �� �� �� �� �� �� �� �� �� Ϣ������(lzhiy0816 design)''''''''
else
session("checked")="no"
session("check")="wrong"
response.Redirect "mdflogin.asp?id="&id&""    ''''''''���벻�ԣ��򷵻ء��� �� �� �� �� �� �� �ڡ����(lzhiy0816 design)'''''''
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
