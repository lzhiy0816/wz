<%
''''''''----====    lzhiy0816 ����վ��     ====----''''''''
''''''''         http://lzhiy0816.vicp.net         ''''''''
''''''''         http://lzhiy0816.3322.org         ''''''''

''''''''----====    lzhiy0816 ���Ͽռ�     ====----''''''''
''''''''      http://free.glzc.net/lzhiy0816       ''''''''
''''''''      http://lzhiy0816.go.nease.net        ''''''''

''''''''====������ʺš�(lzhiy0816 design)====''''''''
%>

<html>
<head>
<title>���������</title>
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
name=request.querystring("name")    ''''''''��ȡ��������û��ύ�ġ�վ������������(lzhiy0816 design)''''''''
xingming=name
dim exec,conn,rs
%>
<!--#include file="conn.asp"-->
<%
exec="select xingming from logo"    ''''''''��Ѱxingming�ֶμ�¼(lzhiy0816 design)''''''''
set rs=server.createobject("adodb.recordset")
rs.open exec,conn,1,1

xingming1=""                        ''''''''�Ƚ�xingming��Ϊ�գ�lzhiy0816 design)''''''''
do while not rs.eof                 ''''''''����ύ���û����ڼ�¼����ڣ���xingming1��Ϊ��¼ֵ��lzhiy0816 design)''''''''
  if xingming=rs("xingming") then
  xingming1=rs("xingming")
  end if
rs.movenext
loop
%>

<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
<%if name="" then %>
    <td align="center" valign="middle">����ѡ���û�������Ϊ��</td>
<%elseif xingming1<>"" then %>
    <td align="center" valign="middle">����ѡ���û���"<font color="#FF0000"><%=xingming1%></font>"�Ѿ����û�ʹ�ã�������ѡ��һ���µ��û�����</td>
<%else %>
    <td align="center" valign="middle">�û���"<font color="#008800"><%=name%></font>"��������ʹ�á�</td>
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
