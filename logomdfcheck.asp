<%
''''''''====����վվ�������顷(lzhiy0816 design)====''''''''
dim inputxm,inputpw    ''''''''��ȡ��վվ���������վ�ó�����(lzhiy0816 design)''''''''
inputxm=request.form("xingming")
inputpw=request.form("password")
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%

exec="select * from admin where id=1"    ''''''''��ȡlzhiy0816��"xingming"��"password"ֵ(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
adminxm=rs("xingming")
adminpw=rs("password")
rs.Close
set rs=nothing

''''''''�����lzhiy0816����ص�logo.asp��lzhhiy0816�������(lzhiy0816 design)''''''''
if inputxm=adminxm and inputpw=adminpw then
response.Redirect "logolzy.asp"
else
exec="select * from logo where xingming='"&inputxm&"'and password='"&inputpw&"'"     ''''''''��Ѱ��վվ�����û���������(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
''''''''�������վվ��������롶�޸����ӹ����ڡ�����(lzhiy0816 design)''''''''
do while not rs.eof 
inputxm1=rs("xingming")
inputpw1=rs("password")
rs.movenext 
loop 
if inputxm1<>"" then
session("checked")="yes2"
session("check")="right2"
response.Redirect "logomdflogin.asp?user="&inputxm1
''''''''���������վվ��,�ص����� վ վ �� �� ֤������(lzhiy0816 design)''''''''
else
session("checked")="no2"
session("check")="wrong2"
response.Redirect "logomdf.asp"
end if
rs.Close
set rs=nothing
end if
conn.close
set conn=nothing
%>
 
