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
''''''''====《管理员提交“回复”或“删除”后的处理程序》(lzhiy0816 design)====''''''''
if "回复"=Request.Form("Submit") then    ''''''''如果提交值为回复(lzhiy0816 design)''''''''
exec="select * from userleave where id="&request.form("id")    ''''''''查寻id号(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,3
hui=request.form("huifu")
huifu=Replace(hui,vbcrlf,"<Br>")
huifu=Replace(huifu," ","&nbsp;")
rs("huifu")=huifu    ''''''''以上是lzhiy0816添加或修改某id号的回复内容(lzhiy0816 design)''''''''
rs.update
rs.close
set rs=nothing
end if

if "删除"=Request.Form("Submit") then    ''''''''如果提交值为删除(lzhiy0816 design)''''''''
exec="delete * from userleave where id="&request.form("id")    ''''''''lzhiy0816删除某id号的留言内容(lzhiy0816 design)''''''''
conn.execute exec
end if

if "显IP"=Request.Form("Submit") then    ''''''''如果提交值为显IP(lzhiy0816 design)''''''''
session("checked")="yes3"
session("check")="right3"
response.Redirect "leaveword.asp"
end if

conn.close
set conn=nothing
response.redirect "leaveword.asp"
%>
