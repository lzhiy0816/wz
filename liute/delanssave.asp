<%
''''''''----====    lzhiy0816 ����վ��     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 ���Ͽռ�     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''
%>

<!--#include file="conn.asp"-->
<%
''''''''====������Ա�ύ���ظ�����ɾ������Ĵ������(lzhiy0816 design)====''''''''
if "�ظ�"=Request.Form("Submit") then    ''''''''����ύֵΪ�ظ�(lzhiy0816 design)''''''''
exec="select * from userleave where id="&request.form("id")    ''''''''��Ѱid��(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
hui=request.form("huifu")
huifu=Replace(hui,vbcrlf,"<Br>")
huifu=Replace(huifu," ","&nbsp;")
rs("huifu")=huifu    ''''''''������lzhiy0816��ӻ��޸�ĳid�ŵĻظ�����(lzhiy0816 design)''''''''
rs.update
rs.close
set rs=nothing
end if

if "ɾ��"=Request.Form("Submit") then    ''''''''����ύֵΪɾ��(lzhiy0816 design)''''''''
exec="delete * from userleave where id="&request.form("id")    ''''''''lzhiy0816ɾ��ĳid�ŵ���������(lzhiy0816 design)''''''''
conn.execute exec
end if

if "��IP"=Request.Form("Submit") then    ''''''''����ύֵΪ��IP(lzhiy0816 design)''''''''
session("checked")="yes3"
session("check")="right3"
response.Redirect "leaveword.asp"
end if

conn.close
set conn=nothing
response.redirect "leaveword.asp"
%>
