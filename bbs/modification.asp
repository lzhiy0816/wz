<!-- #include file="setup.asp" -->
<%
top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<title>�޸�����</title>


<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img
src="images/closedfold.gif" width="14" height="14">��<a href=main.asp><%=clubname%></a><br>
&nbsp;<img
height=15
src="images/coner.gif"
width=15><img
src="images/openfold.gif">����������</font></td>
</tr>
</table>









<br>
<form action="RecoverPasswd.asp" method="post">


<table width="333" border="0" cellspacing="1" cellpadding="2" align="center" class=a2>
<tr>
<td width="328" class=a1 height="20"><div align="center">
<font color=000000>��������</font></div>
</td></tr><tr>
<td height="19" width="328" valign="top" class=a3>
<div align="center">�û����ƣ�<input size="25" name="username" value="<%=Request.Cookies("username")%>"></div>
<div align="center">�������ڣ�<input onfocus="show_cele_date(birthday,'','',birthday)" name="birthday" size="25"><br>
<input type="submit" value=" ȷ�� " name="Submit1">��<input type="reset" value=" ȡ�� " name="Submit">
</div></td></tr> </table></form><center><a href=javascript:history.back()>BACK </a><br>

<%
htmlend
%>