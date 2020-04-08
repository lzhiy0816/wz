<!-- #include file="setup.asp" -->
<%
top
id=int(Request("id"))
forumid=Conn.Execute("Select forumid From forum where id="&id)(0)
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
content=replace(server.htmlencode(Request.Form("content")), "'", "&#39;")
content=replace(content,vbCrlf,"<br>")
content=replace(content,"\","&#92;")


if  Request.ServerVariables("request_method") = "POST" and Request("preview")="" then

If not conn.Execute("Select username From [prison] where username='"&Request.Cookies("username")&"'" ).eof Then error("<li>您被关进<a href=prison.asp>监狱</a>")

username=Request.Cookies("username")


if conn.execute("Select count(id)from forum where ID="&ID&" and forumid="&forumid&" and locktopic=1")(0)=1 then
error2("此主题已经关闭，不接受新的回复")
end if


if content="" then
message=message&"<li>内容没有填写"
end if

''''''''''''''''''''

if Len(content)>50000 then
message=message&"<li>文章内容不能大于 50000 字节"
end if


if Application("content")=content then
message=message&"<li>请不要发表相同的文章"
end if


if message<>"" then
error(""&message&"")
end if


sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3


if rs("userlife")<5 then
message=message&"<li>您的体力值 < <FONT color=red>5</FONT> 不能发表文章<li>您可以到<A href=shop.asp>社区商店</A>购买体力药丸<li>每有效停留时间<FONT color=red> 10 </FONT>分钟：体力值：<FONT color=red>+10</FONT>"
end if

if rs("experience")<1 and Reg10 = 1 then
message=message&"<li>新注册用户登录社区停留<FONT color=red> 10 </FONT>分钟以上才可发表帖子"
end if


if message<>"" then
error(""&message&"")
end if



if rs("membercode")<2 and instr(Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0),"|"&Request.Cookies("username")&"|")=0 then
rs("userlife")=rs("userlife")-2
end if



rs("landtime")=now()
rs("postrevert")=rs("postrevert")+1
rs("money")=rs("money")+2
rs("experience")=rs("experience")+2
rs.update
rs.close

if Request.Cookies("filename")<>empty then
master=split(Request.Cookies("filename"),"|")
for i = 1 to ubound(master)
conn.execute("update [upfile] set id='"&id&"' where filename='"&master(i)&"'")
next
Response.Cookies("filename")=""
end if

sql="insert into reforum (topicid,username,content,postip) values ('"&id&"','"&Request.Cookies("username")&"','"&content&"','"&remoteaddr&"')"
conn.Execute(SQL)

conn.execute("update [forum] set lastname='"&username&"',replies=replies+1,lasttime='"&now()&"' where ID="&ID&" and forumid="&forumid&"")

conn.execute("update [bbsconfig] set lasttime='"&now()&"',lastname='"&username&"',tolrestore=tolrestore+1 where id="&forumid&"")

Application("content") = content

message=message&"<li>回复主题成功<li><a onclick=min_yuzi() target=message href=ShowPost.asp?id="&id&">返回主题</a><li><a href=ShowForum.asp?forumid="&forumid&">返回论坛</a><li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")


end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



sql="select * from bbsconfig where id="&forumid&""
rs.Open sql,Conn
bbsname=rs("bbsname")
logo=rs("logo")
rs.close




%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title>回复文章</title>
<CENTER>




<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<SCRIPT>if ("<%=logo%>"!=''){document.write("<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'")}else{document.write("<img border=0 src=images/lzylogo.gif>")}</SCRIPT>
</td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/closedfold.gif" border="0">　<a href="ShowForum.asp?forumid=<%=forumid%>"><%=bbsname%></a><br>
<%


if Request("quote")=1 and Request("preview")="" then

if isnumeric(""&Request("retopicid")&"") then
sql="select * from reforum where id="&Request("retopicid")&""
else
sql="select * from forum where ID="&id&""
end if
rs.Open sql,Conn
content =rs("content")
quote="[quote]原文由 [B]"&rs("username")&"[/B] 发表:"&vbCrlf&""&content&""&vbCrlf&"[/quote]"
rs.close

end if


%>




　　&nbsp;<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<a onclick=min_yuzi() target=message href="ShowPost.asp?id=<%=id%>">Re:<%=Request("topic")%></a></td>
</tr>
</table>

<br>
<%

if Request("preview")<>"" then
%>
<script src="inc/ybb.js"></script>
<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=5 width=97% border=0 class=a2>
<TR>
<TD vAlign=left class=a1 height=13><b>预览帖子</b></TD></TR>
<TR>
<TD vAlign=left class=a3 height=12>
<script>
var badwords= "<%=badwords%>";
document.write(ybbcode("<%=content%>"));
</script>
</TD></TR>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<br>
<%
end if
%>



<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=5 width=97% border=0 class=a2>



<form method=post name=form onsubmit="return ValidateForm()">
<input type=hidden name=id value=<%=id%>>
<TBODY>
<TR>
<TD vAlign=left class=a1 colSpan=2 height=25><b>回复文章</b></TD></TR>
<TR>
<TD width=160 class=a4 height=25><B>用户名称</B></TD>
<TD width=574 class=a4 height=25><%=Request.Cookies("username")%> [<a href="login.asp?menu=out">退出登录</a>]</TD>
</TR>
<TR>
<TD vAlign=top class=a3 rowSpan=2>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=left border=0>
<TBODY>
<TR>
<TD vAlign=top align=left width=100% class=a3><B>文章内容</B><BR>
(<a href="javascript:CheckLength();">查看内容长度</a>)<BR><BR>
<BR></TD></TR>
<TR>
<TD vAlign=center align=center width=100%>
<TABLE
style="BORDER-RIGHT: 1px inset; BORDER-TOP: 1px inset; BORDER-LEFT: 1px inset; BORDER-BOTTOM: 1px inset"
cellSpacing=1 cellPadding=3 border=0>
<TR align=middle>
<script>
ii=0
for(i=1;i<=25;i=i+1) {
index = Math.floor(Math.random() * 80+1);
ii=ii+1
document.write("<TD><a href=javascript:emoticon('[em"+index+"]')><img border=0 src=images/Emotions/"+index+".gif></a></TD>")
if (ii==5){document.write("</TR><TR align=middle>");ii=0;}
}
</script>
</TR></TABLE>

</TD></TR></TBODY></TABLE></TD>
<SCRIPT src="inc/post.js"></SCRIPT>
<SCRIPT src="inc/ybbcode.js"></SCRIPT>
<TD width=574 class=a3>

<img border=0 src=images\ybb\bold.gif onclick=YBBbold() alt="粗体">
<img border=0 src=images\ybb\italicize.gif onclick=YBBitalic() alt="斜体">
<img border=0 src=images\ybb\underline.gif onclick=YBBunder() alt="下划线">
<img border=0 src=images\ybb\center.gif onclick=YBBcenter() alt="居中">
<img border=0 src=images\ybb\url.gif onclick=YBBurl() alt="插入超链接">
<img border=0 src=images\ybb\email.gif onclick=YBBemail() alt="插入电子邮件地址">
<img border=0 src=images\ybb\image.gif onclick=YBBimage() alt="插入图片">
<img border=0 src=images\ybb\flash.gif onclick=YBBflash() alt="插入FLASH文件">
<img border=0 src=images\ybb\rm.gif onclick=YBBrm() alt="插入RealPlayer文件">
<img border=0 src=images\ybb\mp.gif onclick=YBBmp() alt="插入Media Player文件">
<img border=0 src=images\ybb\quote.gif onclick=YBBquote() alt="引用">
<select onchange=ybbfont(this.options[this.selectedIndex].value)>
<option>字体</option>
<option value="Arial">Arial</option>
<option value="Arial Black">Arial Black</option>
<option value="Verdana">Verdana</option>
<option value="Times New Roman">Times New Roman</option>
<option value="Garamond">Garamond</option>
<option value="Courier New">Courier New</option>
<option value="Webdings">Webdings</option>
<option value="Wingdings">Wingdings</option>
<option value="隶书">隶书</option>
<option value="幼圆">幼圆</option>
<OPTION value="方正舒体">方正舒体</OPTION>
<OPTION value="方正姚体">方正姚体</OPTION>
<OPTION value="仿宋_GB2312">仿宋_GB2312</OPTION>
<OPTION value="黑体">黑体</OPTION>
<OPTION value="华文彩云">华文彩云</OPTION>
<OPTION value="华文细黑">华文细黑</OPTION>
<OPTION value="华文新魏">华文新魏</OPTION>
<OPTION value="华文行楷">华文行楷</OPTION>
<OPTION value="华文中宋">华文中宋</OPTION>
<OPTION value="楷体_GB2312">楷体_GB2312</OPTION>
<OPTION value="隶书">隶书</OPTION>
<OPTION value="宋体">宋体</OPTION>
<OPTION value="新宋体">新宋体</OPTION>
<OPTION value="幼圆">幼圆</OPTION>
</SELECT>
<select onchange=ybbsize(this.options[this.selectedIndex].value)>
<OPTION>大小</OPTION>
<OPTION value=1>1</OPTION>
<OPTION value=2>2</OPTION>
<OPTION value=3>3</OPTION>
<OPTION value=4>4</OPTION>
<OPTION value=5>5</OPTION>
<OPTION value=6>6</OPTION>
<OPTION value=7>7</OPTION>
</SELECT>

<TABLE cellSpacing=0 cellPadding=0 align=left border=0 >
	<TBODY>
	<TR><TD id=ColorUsed onclick="if(this.bgColor.length > 0) insertTag(this.bgColor)" vAlign=center align=middle>
	<IMG height=10 width=20 border=1></TD></TD><SCRIPT>rgb(65,1,10)</SCRIPT></TR>
	</TBODY>
</TABLE>

</TD></TR>
<TR>
<TD width=574 class=a3>
<div id="sending" style="position:absolute;visibility:hidden;"><table width=520 height=170 border=0 cellspacing=2 cellpadding=0 class=a2><tr><td class=a3 align=center>内容正在发送, 请稍候...</td></tr></table></div>

<TEXTAREA onkeydown=presskey(); name=content rows=10 cols=70><%=quote%><%=Request.Form("content")%></TEXTAREA>
<br>　　　　　　　　　　　　　『 <a href="javascript:replac()">替换内容</a> 』 『  
<a href="javascript:HighlightAll('form.content')">复制到剪贴板</a> 』</TD></TR>



<TR>
<TD align=left class=a4>
<IMG src=images/affix.gif><b>增加附件</b></TD>

<TD align=left class=a4><font color="FFFFFF"><b><IFRAME src="upfile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></b></font></TD></TR>

<TR>
<TD align=middle class=a3 colSpan=2 height=27><INPUT tabIndex=4 type=submit value=回复主题 name=submit1>&nbsp; <input name=preview type="submit" value=" 预 览 ">   <INPUT type=reset onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" value=" 重 写 "></TD></TR></TBODY></FORM>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT>

<Script>
document.form.content.value = unybbcode(document.form.content.value);
document.form.content.focus();
function unybbcode(temp) {
temp = temp.replace(/\[quote](.*)\[\/quote]/gi,"");
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>

<%

htmlend
%>