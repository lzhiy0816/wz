<!-- #include file="setup.asp" -->
<%

if apply=0 then error("<li>��������ʱֹͣ��̳����")

top

if Request("menu")="apply" then

if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")



if Request("bbsname")="" then error("<li>��������̳����")

if Len(Request("intro"))>255 then  error("<li>��̳��鲻�ܴ��� 255 ���ֽ�")


if instr(Application("apply"),""&Request.Cookies("username")&" ")>0 then error("<li>���Ѿ��������̳�ˣ��벻Ҫ�ٴ����룡")




rs.Open "bbsconfig",Conn,1,3

rs.addnew
rs("bbsname")=server.htmlencode(Request("bbsname"))
rs("moderated")="|"&Request.Cookies("username")&"|"
rs("intro")=server.htmlencode(Request("intro"))
rs("icon")=server.htmlencode(Request("icon"))
rs("logo")=server.htmlencode(Request("logo"))
rs("hide")=1
rs.update
id=rs("id")
rs.close

mailaddress=Conn.Execute("Select usermail From [user] where username='"&Request.Cookies("username")&"'")(0)
mailtopic="��̳ϵͳ��֪ͨͨ"
body=""&vbCrlf&"�װ���"&Request.Cookies("username")&", ����!"&vbCrlf&""&vbCrlf&"������ϲ! ���Ѿ��ɹ���������"&homename&"("&homeurl&")����̳ϵͳ, �ǳ���л��ʹ��"&homename&"�ķ���!"&vbCrlf&""&vbCrlf&"��* �������Ϊ������̳�ṩ��һ���ȽϺüǵĵ�ַ,��������"&vbCrlf&""&vbCrlf&"��* URL: "&Request("bbsname")&"("&cluburl&"/Default.asp?id="&id&")"&vbCrlf&""&vbCrlf&"��* ���, �м���ע�����������μ�"&vbCrlf&"1������ʹ�ñ���̳ϵͳ�����κΰ���ɫ�顢�Ƿ����Լ�Σ�����Ұ�ȫ�����ݵ���̳��"&vbCrlf&"2�������ڱ�ϵͳ�û���ӵ�е���̳�ڷ����κ�ɫ�顢�Ƿ�������Σ�����Ұ�ȫ�����ۡ�"&vbCrlf&"3�����Ϲ���Υ�������Ը�����վ��Ȩɾ�������û��������ݣ���׷���䷨�����Ρ�"&vbCrlf&""&vbCrlf&""&vbCrlf&"��̳������ "&homename&"("&homeurl&") �ṩ������������Yuzi������(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/mail.asp" -->
<%


Application("apply")=""&Request.Cookies("username")&" "&Application("apply")&""





message="<li>��ϲ��������ɹ���<li>����Ϊ��������̳����ṩ��һ���������Լ�������<li><a target=_blank href=Default.asp?id="&id&">"&cluburl&"/Default.asp?id="&id&"</a><li> ���� <a target=_blank href=Default.asp?id="&id&">"&Request("bbsname")&"</a> �ϵ���������Ҽ����������������������ǩ�����ղؼ���"
succeed(""&message&"")


end if


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
 <title>������̳</title><table width=97% align="center" border="0"><tr><td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td><td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>��<img src="images/bar.gif" border="0"><img src="images/openfold.gif">��������̳</td></tr></table>
<br><center>
<form name="form" method="post" action="apply.asp">
<input type=hidden name=menu value="apply">

<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>������̳</b></td>
</tr>
<tr bgcolor="FFFFFF">
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;<img src="images/1.gif">��̳��Ϣ</b></td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="20%">��̳����<b>��</b></td>
<td class=a3 height="5" align="left" valign="middle" width="78%">
&nbsp;
<input type="text" name="bbsname" size="30" maxlength="12">
</td>
</tr>

<tr class=a4>
<td height="2" align="right" width="20%">��̳����<b>��</b></td>
<td height="2" align="left" valign="middle" width="78%" class=a3>
&nbsp;
<textarea name="intro" rows="4" cols="50"></textarea>&nbsp;
</td>
</tr>
<tr class=a3>
<td height="1" align="right" valign="middle" width="20%">Сͼ��URL��</td>
<td height="1" align="left" valign="middle" width="78%">
&nbsp;
<input size="30" name="icon">����ʾ����̳�������</td>
</tr>
<tr class=a4>
<td align="right" valign="bottom" width="20%">��ͼ��URL��</td>
<td align="left" valign="bottom" width="78%">
&nbsp;
<input size="30" name="logo">����ʾ����̳���Ͻ�</td>
</tr>
<tr>
<td align="right" width="98%" colspan="2" bgcolor="#FFFFFF">
<input type="submit" value=" ȷ �� &gt;&gt;�� һ �� " name="Submit"> </td>
</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
</form>
<%
htmlend
%>