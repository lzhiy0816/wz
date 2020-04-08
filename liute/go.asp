<%
''''''''----====    lzhiy0816 本地站点     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 网上空间     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''
%>

<!--#include file="conn.asp"-->
<%
''''''''====《点击“转至”按钮后》(lzhiy0816 design)====''''''''
exec="select * from userleave Order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1      ''''''''读取记录''''''''
rs.PageSize=10             ''''''''定义每页的记录数''''''''
pagecount=rs.PageCount     ''''''''记录总数''''''''
page=int(request.form("page"))
if page>pagecount then page=pagecount    ''''''''获取页数超出记录总数，则页数等于总数''''''''
if page<1 then page=1                    ''''''''获取页数小于1，则页数等于1''''''''
response.Redirect "leaveword.asp?page="&page&""
%>
