<!-- #include file="setup.asp" -->
<%
if ""&selectmail&""="" then
error2("对不起，本系统尚未开通此功能")
end if
if Request("menu")="ok" then
mailaddress=Request("Friend_email")
mailtopic=""&Request("title")&""

body=""&vbCrlf&""&Request("Friend_name")&"，您好！"&vbCrlf&""&vbCrlf&"我想您可能会喜欢看这内容的，所以我发给您这网页的链接"&vbCrlf&""&vbCrlf&""&Request("url")&""&vbCrlf&""&vbCrlf&"============================================================"&vbCrlf&""
body = body&""&Request("comment")&""
body = body&""&vbCrlf&"============================================================"&vbCrlf&"　　　　　　　　　　　　　　　　　　　　您的朋友："&Request("username")&""&vbCrlf&"　　　　　　　　　　　　　　　　　　　　　　"&Request("mail")&""&vbCrlf&""&vbCrlf&"论坛服务由 "&homename&"("&homeurl&") 提供　程序制作：Yuzi工作室(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""

%>
<!-- #include file="inc/mail.asp" -->
<HTML><HEAD><meta http-equiv=Content-Type content=text/html; charset=gb2312><link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet></HEAD>
<body topmargin=0>
<title>Send to a friend</title><center><TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="458">
<TR><TD VALIGN="TOP" COLSPAN="4" WIDTH="458"><TABLE BORDER="0" CELLPADDING="0" CELLSPACING="0" WIDTH="100%">
<TR><TD VALIGN="TOP" WIDTH="12">　</TD>
<TD WIDTH="100%" class=a1 HEIGHT="30" COLSPAN="3">　<b>谢谢使用</b></TD>
<TD VALIGN="BOTTOM" WIDTH="71"><IMG SRC="images/email_icon01.gif" HEIGHT="30" WIDTH="70"></TD>
<TD VALIGN="TOP" WIDTH="18" class=a1>　&nbsp;</TD>
</TR><TR><TD VALIGN="TOP" WIDTH="12">　</TD><TD NOWRAP VALIGN="BOTTOM" WIDTH="370" COLSPAN="3"></TD>
<TD VALIGN="TOP" WIDTH="71"><IMG SRC="images/email_icon02.gif" HEIGHT="31" WIDTH="70"></TD>
<TD VALIGN="TOP" WIDTH="18">　</TD></TR></TABLE></TD></TR><TR><TD COLSPAN="2"><IMG SRC="images/Trans.gif" HEIGHT="60" WIDTH="1" BORDER="0"></TD></TR>
<TR><TD VALIGN="BOTTOM" ALIGN="LEFT"><IMG SRC="images/Trans.gif" HEIGHT="1" WIDTH="140" BORDER="0"><b>您的信息已经被发送致</b><BR><IMG SRC="images/Trans.gif" HEIGHT="1" WIDTH="140" BORDER="0"><%=Request("Friend_email")%></FONT></TD></TR>
<TR><TD><IMG SRC="images/Trans.gif" HEIGHT="77" WIDTH="280" BORDER="0"></TD><TD VALIGN="BOTTOM"><INPUT TYPE="BUTTON" STYLE="font-family:宋体;font-size:9pt" VALUE=" 关闭 " OnClick="window.close();" ID="button1" NAME="button1" TITLE="Close"></TD></TR>
<TR><TD><IMG SRC="images/Trans.gif" HEIGHT="7" WIDTH="1" BORDER="0"></TD></TR><TR><TD HEIGHT="7" class=a1 COLSPAN="3"></TD></TR>
</TABLE></html>
<%
responseend
end if

%>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<body topmargin=0>
<TITLE>Send to a friend</TITLE>
<CENTER>
<TABLE cellSpacing=0 cellPadding=0 width=458 border=0>
<TBODY>
<TR>
<TD vAlign=top width=458 colSpan=4>
<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD vAlign=top width=12></TD>
<TD width="100%" colSpan=3 height=30 class=a1>　<B>将此页发给您的朋友</B></TD>
<TD vAlign=bottom width=71><IMG height=30 src="images/email_icon01.gif" width=70></TD>
<TD vAlign=top width=18 class=a1>　&nbsp;</TD></TR>
<TR>
<TD vAlign=top width=12>　</TD>
<TD vAlign=bottom noWrap width=370 colSpan=3>如果您认为您的朋友同样对这篇文章感兴趣的话，您只需填写几个简单的信息后点“发送”即可将此文的地址发送给您的朋友</TD>
<TD vAlign=top width=71><IMG height=31 src="images/email_icon02.gif" width=70></TD>
<TD vAlign=top width=18>　</TD></TR></TBODY></TABLE></TD></TR>
<TR>
<TD width=12 rowSpan=3>　</TD></TR>
<TR>

<FORM method=post>
<input type="hidden" name="menu" value="ok">
<input type="hidden" name="url" value="<%=Request.ServerVariables("http_referer")%>">
<input type="hidden" name="title" value="<%=Request("title")%>">

<TD vAlign=top width=270><FONT color=3366cc>注释（可选项）</FONT> <TEXTAREA name=comment rows=8 wrap=hard cols=35></TEXTAREA><BR clear=all>请您填写姓名及E-mail只是为了您的朋友方便的了解此条信息的来源！</FONT></TD>
<TD width=12>　</TD>
<TD vAlign=top width=185><FONT color=3366cc>您朋友的姓名</FONT>
<INPUT name=Friend_name size="20"><BR><FONT color=3366cc>您朋友的E-mail地址</FONT>
<INPUT name=Friend_email size="20"><BR><FONT color=3366cc>您的姓名</FONT><BR>
<INPUT name=username size="20" value="<%=Request.Cookies("username")%>"><BR><FONT color=3366cc>您的E-mail地址</FONT>
<INPUT name=mail size="20"><BR>&nbsp;　<NOBR><INPUT title=Send style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体" type=submit value=" 发送 " name=CheckSubmit>　<INPUT title=Cancel style="FONT-SIZE: 9pt; FONT-FAMILY: 宋体" type=reset value=" 取消 " OnClick="window.close();">
</NOBR></TD></TR></FORM>
<TR>
<TD></TD></TR>
<TR>
<TD class=a1 colSpan=4 height=7></TD></TR></TBODY></TABLE></CENTER></BODY></HTML>