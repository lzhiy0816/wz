<%
''''''''====������Ա������logo��(lzhiy0816 design)====''''''''
dim inputxm,inputpw    ''''''''��ȡlzhiy0816����(lzhiy0816 design)''''''''
inputxm=request.form("xingming")
inputpw=request.form("password")
id=request.querystring("id")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select * from admin where id=1"    ''''''''��Ѱlzhiy0816��id��(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
''''''''�����lzhiy0816����롶�� վ �� �� Ա �� �� �� �ڡ�����(lzhiy0816 design)''''''''
if inputxm=rs("xingming")  and inputpw=rs("password") and rs("password")<>"" then
session("lzychecked")="yes"
session("lzycheck")="right"
response.Redirect "logolzymdflogin.asp?id="&id&""
''''''''�������lzhiy0816�ص����� վ �� �� Ա �� �ڡ�����(lzhiy0816 design)''''''''
else
session("lzychecked")="no"
session("lzycheck")="wrong"
response.Redirect "logolzymdf.asp?id="&id&""
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
