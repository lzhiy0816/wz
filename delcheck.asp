<%
''''''''====������Ա�����顷(lzhiy0816 design)====''''''''
dim inputxm2,inputpw2    ''''''''��ȡlzhiy0816����(lzhiy0816 design)''''''''
inputxm2=request.form("xingming2")
inputpw2=request.form("password2")
id=request.querystring("id")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select * from admin where id=1"    ''''''''��Ѱlzhiy0816��id��(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn
''''''''�����lzhiy0816����롶�� վ �� �� Ա �� �� �� �ڡ�����(lzhiy0816 design)''''''''
if inputxm2=rs("xingming")  and inputpw2=rs("password") and rs("password")<>"" then
session("checked")="yes2"
session("check")="right2"
response.Redirect "delorans.asp?id="&id&""
''''''''�������lzhiy0816�ص����� վ �� �� Ա �� �ڡ�����(lzhiy0816 design)''''''''
else
session("checked")="no2"
session("check")="wrong2"
response.Redirect "dellogin.asp?id="&id&""
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
 
