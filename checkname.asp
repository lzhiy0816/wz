<%
''''''''----====    lzhiy0816 本地站点     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 网上空间     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''

''''''''====《检测帐号》(lzhiy0816 design)====''''''''
%>

<html>
<head>
<title>共享的世界</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<style type=text/css>BODY {
	FONT-SIZE: 12px; COLOR: #000000; BACKGROUND-COLOR: #ffffff
}
TD {
	FONT-SIZE: 12px; COLOR: #000000
}
TR {
	FONT-SIZE: 12px; COLOR: #000000
}
A:link {
	COLOR: #0000cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #990000; TEXT-DECORATION: none
}
A:hover {
	COLOR: none; text-decoration: underline
}
FONT {
	TEXT-DECORATION: none
}
INPUT.text {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
INPUT.file {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
SELECT {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
TEXTAREA {
	BORDER-RIGHT: #101010 1px solid; BORDER-TOP: #101010 1px solid; BORDER-LEFT: #101010 1px solid; COLOR: #000000; BORDER-BOTTOM: #101010 1px solid; BACKGROUND-COLOR: #ffffff
}
INPUT.buttom {
	BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 12px; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; COLOR: #000000; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BACKGROUND-COLOR: #dedfdf; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000
}
</style>

<body bgcolor="#FFFFFF" text="#000000" background="img/background/lzybg01.jpg">
<%
name=request.querystring("name")    ''''''''获取申请加入用户提交的“站长姓名”数据(lzhiy0816 design)''''''''
xingming=name
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select xingming from logo"    ''''''''查寻xingming字段记录(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1

xingming1=""                        ''''''''先将xingming置为空（lzhiy0816 design)''''''''
do while not rs.eof                 ''''''''如果提交的用户名在记录里存在，则将xingming1置为记录值（lzhiy0816 design)''''''''
  if xingming=rs("xingming") then
  xingming1=rs("xingming")
  end if
rs.movenext
loop
%>

<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
<%if name="" then %>
    <td align="center" valign="middle">您所选的用户名不能为空</td>
<%elseif xingming1<>"" then %>
    <td align="center" valign="middle">您所选的用户名"<font color="#FF0000"><%=xingming1%></font>"已经有用户使用，请另外选择一个新的用户名。</td>
<%else %>
    <td align="center" valign="middle">用户名"<font color="#008800"><%=name%></font>"可以正常使用。</td>
<%end if %>
  </tr>
</table>
<%
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>
