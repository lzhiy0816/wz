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
''''''''====�������ת������ť��(lzhiy0816 design)====''''''''
exec="select * from userleave Order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1      ''''''''��ȡ��¼''''''''
rs.PageSize=10             ''''''''����ÿҳ�ļ�¼��''''''''
pagecount=rs.PageCount     ''''''''��¼����''''''''
page=int(request.form("page"))
if page>pagecount then page=pagecount    ''''''''��ȡҳ��������¼��������ҳ����������''''''''
if page<1 then page=1                    ''''''''��ȡҳ��С��1����ҳ������1''''''''
response.Redirect "leaveword.asp?page="&page&""
%>
