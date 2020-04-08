<!-- #include file="setup.asp" -->
<%

if apply=0 then error("<li>本社区暂时停止论坛申请")

top

if Request("menu")="apply" then

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")



if Request("bbsname")="" then error("<li>请输入论坛名称")

if Len(Request("intro"))>255 then  error("<li>论坛简介不能大于 255 个字节")


if instr(Application("apply"),""&Request.Cookies("username")&" ")>0 then error("<li>您已经申请过论坛了，请不要再次申请！")




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
mailtopic="论坛系统开通通知"
body=""&vbCrlf&"亲爱的"&Request.Cookies("username")&", 您好!"&vbCrlf&""&vbCrlf&"　　恭喜! 您已经成功地申请了"&homename&"("&homeurl&")的论坛系统, 非常感谢您使用"&homename&"的服务!"&vbCrlf&""&vbCrlf&"　* 我们免费为您的论坛提供了一个比较好记的地址,请您试试"&vbCrlf&""&vbCrlf&"　* URL: "&Request("bbsname")&"("&cluburl&"/Default.asp?id="&id&")"&vbCrlf&""&vbCrlf&"　* 最后, 有几点注意事项请您牢记"&vbCrlf&"1、不得使用本论坛系统建立任何包含色情、非法、以及危害国家安全的内容的论坛。"&vbCrlf&"2、不得在本系统用户所拥有的论坛内发布任何色情、非法、或者危害国家安全的言论。"&vbCrlf&"3、以上规则违者责任自负，本站有权删除该类用户或者内容，并追究其法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&homename&"("&homeurl&") 提供　程序制作：Yuzi工作室(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/mail.asp" -->
<%


Application("apply")=""&Request.Cookies("username")&" "&Application("apply")&""





message="<li>恭喜您，申请成功！<li>我们为您的新论坛免费提供了一个属于您自己的域名<li><a target=_blank href=Default.asp?id="&id&">"&cluburl&"/Default.asp?id="&id&"</a><li> 请在 <a target=_blank href=Default.asp?id="&id&">"&Request("bbsname")&"</a> 上点击您鼠标的右键，把这个域名加入您的书签或者收藏夹中"
succeed(""&message&"")


end if


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
 <title>申请论坛</title><table width=97% align="center" border="0"><tr><td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td><td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>　<img src="images/bar.gif" border="0"><img src="images/openfold.gif">　申请论坛</td></tr></table>
<br><center>
<form name="form" method="post" action="apply.asp">
<input type=hidden name=menu value="apply">

<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>申请论坛</b></td>
</tr>
<tr bgcolor="FFFFFF">
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;<img src="images/1.gif">论坛信息</b></td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="20%">论坛名称<b>：</b></td>
<td class=a3 height="5" align="left" valign="middle" width="78%">
&nbsp;
<input type="text" name="bbsname" size="30" maxlength="12">
</td>
</tr>

<tr class=a4>
<td height="2" align="right" width="20%">论坛介绍<b>：</b></td>
<td height="2" align="left" valign="middle" width="78%" class=a3>
&nbsp;
<textarea name="intro" rows="4" cols="50"></textarea>&nbsp;
</td>
</tr>
<tr class=a3>
<td height="1" align="right" valign="middle" width="20%">小图标URL：</td>
<td height="1" align="left" valign="middle" width="78%">
&nbsp;
<input size="30" name="icon">　显示在论坛介绍左边</td>
</tr>
<tr class=a4>
<td align="right" valign="bottom" width="20%">大图标URL：</td>
<td align="left" valign="bottom" width="78%">
&nbsp;
<input size="30" name="logo">　显示在论坛左上角</td>
</tr>
<tr>
<td align="right" width="98%" colspan="2" bgcolor="#FFFFFF">
<input type="submit" value=" 确 定 &gt;&gt;下 一 步 " name="Submit"> </td>
</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
</form>
<%
htmlend
%>