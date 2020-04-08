<!--#include file="conn.asp"-->
<%

exec="select * from admin where id=1"    ''''''''获取lzhiy0816的"face"值(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1
PageSize=rs("PageSizelogo")
rs.Close
set rs=nothing


''''''''====《点击“转至”按钮后(logolzy)》(lzhiy0816 design)====''''''''
exec="select * from logo Order by id desc"
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1      ''''''''读取记录''''''''
rs.PageSize=PageSize             ''''''''定义每页的记录数''''''''
pagecount=rs.PageCount     ''''''''记录总数''''''''
page=int(request.form("page"))
if page>pagecount then page=pagecount    ''''''''获取页数超出记录总数，则页数等于总数''''''''
if page<1 then page=1                    ''''''''获取页数小于1，则页数等于1''''''''
response.Redirect "logolzy.asp?page="&page&""

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
