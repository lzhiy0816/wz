<!--#include file="conn.asp"-->
<%
''''''''====�����û��ύ�޸ĵ��������ݣ���д�����ݿ⡷(lzhiy0816 design)====''''''''

id=request.querystring("id")
password=request.form("password")

if "�ύ"=Request.Form("Submit") then    ''''''''����ύֵΪ"�ύ"(lzhiy0816 design)''''''''
exec="select * from logo where id="&id    ''''''''��Ѱid���û��ύ����������(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3

''''''''�޸��������ݣ���д�����ݿ�(lzhiy0816 design)''''''''
rs("xingming")=request.form("xingming")
if password<>"" then
rs("password")=request.form("password")
end if
rs("mingchen")=request.form("mingchen")
rs("shuoming")=request.form("shuoming")
rs("lianjie")=request.form("lianjie")
rs("logoimg")=request.form("logoimg")

rs.update
rs.close
set rs=nothing
end if

if "ɾ��"=Request.Form("Submit") then    ''''''''����ύֵΪɾ��(lzhiy0816 design)''''''''
exec="delete * from logo where id="&id    ''''''''lzhiy0816ɾ��ĳid�ŵ���������(lzhiy0816 design)''''''''
conn.execute exec
end if

conn.close
set conn=nothing
response.redirect "logo.asp"
%>
