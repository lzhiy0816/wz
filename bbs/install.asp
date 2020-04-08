<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->
<%
youip="127.0.0.1" '您本机的IP地址


if adminpassword<>"" and remoteaddr<>youip then
error("<li>为了安全起见，请编辑 <font color=red>install.asp</font> 内本机的IP地址<li>请把它设置成 <font color=red>"&remoteaddr&"</font>")
end if

administrators=Request("administrators")
if Request("menu")="ok" then

if administrators="" then error("<li>您没有设置社区区长")
if Request("adminpassword")="" then error("<li>您没有设置管理密码")
If conn.Execute("Select id From [user] where username='"&administrators&"'" ).eof Then error("<li>"&administrators&"的用户资料还未<a target=_blank href=register.asp?username="&administrators&">注册</a>")

conn.execute("update [clubconfig] set adminpassword='"&MD5(Request("adminpassword"))&"'")
conn.execute("update [user] set membercode=5 where username='"&administrators&"'")

message=message&"<li>安装成功<li><a href=admin.asp target=_top>返回社区管理</a><li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=admin.asp>")
end if
top
%>
<HTML><META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<title>安装BBSxp...</title><br><form method="post"><input type="hidden" value="ok" name="menu">  </div>
<center><table cellSpacing="0" cellPadding="3" width="160" border="0">
<tr>
<td width="626" bordercolor="FFFFFF">
<table width=435 cellpadding="1" cellspacing="1" border="0" class=a2><tbody>
<tr>
<td align=center colspan="2" height="25" class=a1><p><b>设置社区管理密码</b></p></td>
</tr>


<tr class=a3><td align=right width="30%" height="20">社区区长：</td>
<td height="20">
<input name=administrators size="30"></td></tr>
<tr class=a4><td align=right width="30%" height="10">管理密码：</td>
<td height="10">
<input name=adminpassword type=password size="30"></td></tr>
<tr class=a3><td align=center height="10" colspan="2"><input type="submit" value=" 下一步 "></td>
</tr>
</tbody></table>
</td></tr></table></FORM>
注：<a href="register.asp">如果区长用户名还没有注册请点击这里进行注册</a>
<%htmlend%>