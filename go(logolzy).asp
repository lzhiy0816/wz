<!--#include file="conn.asp"-->
<%

exec="select * from admin where id=1"    ''''''''��ȡlzhiy0816��"face"ֵ(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
PageSize=rs("PageSizelogo")
rs.Close
set rs=nothing


''''''''====�������ת������ť��(logolzy)��(lzhiy0816 design)====''''''''
exec="select * from logo Order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1      ''''''''��ȡ��¼''''''''
rs.PageSize=PageSize             ''''''''����ÿҳ�ļ�¼��''''''''
pagecount=rs.PageCount     ''''''''��¼����''''''''
page=int(request.form("page"))
if page>pagecount then page=pagecount    ''''''''��ȡҳ��������¼��������ҳ����������''''''''
if page<1 then page=1                    ''''''''��ȡҳ��С��1����ҳ������1''''''''
response.Redirect "logolzy.asp?page="&page&""

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
