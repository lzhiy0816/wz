<!-- #include file="setup.asp" -->

<%

top

forumid=int(Request("forumid"))
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


content=server.htmlencode(Request.Form("content"))
content=replace(content,vbCrlf,"<br>")
content=replace(content,"\","&#92;")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" and Request("preview")="" then

If not conn.Execute("Select username From [prison] where username='"&Request.Cookies("username")&"'" ).eof Then error("<li>您被关进<a href=prison.asp>监狱</a>")

username=Request.Cookies("username")
icon=Request("icon")
topic=server.htmlencode(Trim(Request("topic")))
topic=replace(topic,"\","&#92;")
if topic="" then
message=message&"<li>主题没有填写"
end if
if content="" then
message=message&"<li>内容没有填写"
end if

if Len(topic)>50 then
message=message&"<li>文章主题不能大于 50 字节"
end if

if Len(content)>50000 then
message=message&"<li>文章内容不能大于 50000 字节"
end if

if instr(topic,"	") > 0 then
message=message&"<li>主题中不能含有特殊字符"
end if

if Application("content")=content then
message=message&"<li>请不要发表相同的文章"
end if


''''''''''''''''''''


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
rs("userlife")=rs("userlife")-5
end if

rs("landtime")=now()
rs("posttopic")=rs("posttopic")+1
rs("money")=rs("money")+5
rs("experience")=rs("experience")+5
rs.update
rs.close


rs.Open "forum",conn,1,3
rs.addnew
rs("username")=username
rs("lastname")=username
rs("forumid")=forumid
rs("topic")=topic
rs("content")=content
rs("postip")=remoteaddr
rs("icon")=icon

'''''''''''''''''''''''''''''''''''
'投票处理程序
if Request("vote")<>"" then
vote=server.htmlencode(Trim(Request("vote")))
if instr(vote,"|") > 0 then
message=message&"<li>投票选项中不能含有“|”字符"
error(""&message&"")
end if
polltopic=split(vote,chr(13)&chr(10))
j=0
for i = 0 to ubound(polltopic)
if not (polltopic(i)="" or polltopic(i)=" ") then
allpolltopic=""&allpolltopic&""&polltopic(i)&"|"
j=j+1
end if
next
for y = 1 to j
votenum=""&votenum&"0|"
next
rs("polltopic")=allpolltopic
rs("pollresult")=votenum
rs("multiplicity")=Request("multiplicity")
end if
'''''''''''''''''''''''''''''''''''
rs.update
id=rs("id")
rs.close




if Request.Cookies("filename")<>empty then
master=split(Request.Cookies("filename"),"|")
for i = 1 to ubound(master)
conn.execute("update [upfile] set id='"&id&"',topic='"&topic&"' where filename='"&master(i)&"'")
next
Response.Cookies("filename")=""
end if


conn.execute("update [bbsconfig] set lastname='"&username&"',lasttime='"&now()&"',toltopic=toltopic+1,tolrestore=tolrestore+1 where id="&forumid&"")

Application("content") = content

message="<li>新主题发表成功<li><a onclick=min_yuzi() target=message href=ShowPost.asp?id="&id&">返回主题</a><li><a href=ShowForum.asp?forumid="&forumid&">返回论坛</a><li><a href=main.asp>返回论坛首页</a>"
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
<title>发表文章</title>

<CENTER>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<SCRIPT>if ("<%=logo%>"!=''){document.write("<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'")}else{document.write("<img border=0 src=images/lzylogo.gif>")}</SCRIPT>
</td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/closedfold.gif" border="0">　<a href="ShowForum.asp?forumid=<%=forumid%>"><%=bbsname%></a><br>
　　&nbsp;<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　发表文章</td>
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

<TABLE cellSpacing=1 cellPadding=6 width=97% border=0 class=a2>



<form method=post name=form onsubmit="return ValidateForm()">
<input type=hidden name=forumid value=<%=forumid%>>
<TBODY>
<TR>
<TD vAlign=left colSpan=2 height=25 class=a1><b>发表文章</b></TD></TR>
<TR>
<TD width=160 class=a4 height=25><B>用户名称</B></TD>
<TD width=570 class=a4 height=25><%=Request.Cookies("username")%> [<a href="login.asp?menu=out">退出登录</a>]</TD>
</TR>
<TR>
<TD width=160 class=a3 height=25><B>文章标题 </B> <SELECT onchange=DoTitle(this.options[this.selectedIndex].value)>
<OPTION value="" selected>&nbsp;类型</OPTION> <OPTION
value=[原创]>[原创]</OPTION><OPTION value=[转贴]>[转贴]</OPTION> <OPTION
value=[灌水]>[灌水]</OPTION><OPTION value=[讨论]>[讨论]</OPTION> <OPTION
value=[求助]>[求助]</OPTION><OPTION value=[推荐]>[推荐]</OPTION> <OPTION
value=[公告]>[公告]</OPTION><OPTION value=[注意]>[注意]</OPTION> <OPTION
value=[贴图]>[贴图]</OPTION><OPTION value=[建议]>[建议]</OPTION> <OPTION
value=[下载]>[下载]</OPTION><OPTION value=[分享]>[分享]</OPTION></SELECT> </TD>
<TD width=570 class=a3 height=25>
<INPUT maxLength=30 size=60 name=topic value="<%=Request("topic")%>"></TD></TR>
<TR>
<TD vAlign=top align=left width=160 class=a4 height=23><B>您的表情</B></TD>
<TD width=570 class=a4>
<script>
index = Math.floor(Math.random() * 24+1);
for(i=1;i<=24;i=i+1) {
if (i==index) ii="checked";else ii=""
document.write("<INPUT type=radio value="+i+" name=icon "+ii+"><IMG src=images/brow/"+i+".gif>　")
if (i==12) document.write("<br>")
}
</script>

 </TD></TR>

<TR>
<TD vAlign=top class=a3 rowSpan=2>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=left border=0>
<TBODY>
<TR>
<TD vAlign=top align=left width=100% class=a3><B>文章内容</B><BR>
(<a href="javascript:CheckLength();">查看内容长度</a>)<BR>
<BR></TD></TR>
<TR>
<TD align=center width=100%>
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
</TR>

</TABLE>


</TD></TR>

<TR><TD><br>
<INPUT id=advcheck name=advshow type=checkbox value=1 onclick=showadv()><label for=advcheck> 显示投票选项</label>

</TD></TR>

</TBODY></TABLE></TD>
<SCRIPT src="inc/post.js"></SCRIPT>
<SCRIPT src="inc/ybbcode.js"></SCRIPT>
<TD class=a3>
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
</SELECT> <select onchange=ybbsize(this.options[this.selectedIndex].value)>
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
	<IMG height=10 width=20 border=1></TD></TD><SCRIPT>rgb(68,1,10)</SCRIPT></TR>
	</TBODY>
</TABLE>

</TD></TR>


<TR>
<TD width=570 class=a3>

<div id="sending" style="position:absolute;visibility:hidden;"><table width=520 height=170 border=0 cellspacing=2 cellpadding=0 class=a2><tr><td class=a3 align=center>内容正在发送, 请稍候...</td></tr></table></div>

<TEXTAREA onkeydown="presskey();" name=content rows=10 cols=70><%=Request.Form("content")%></TEXTAREA>
<br>
　　　　　　　　　　　　　『 <a href="javascript:replac()">替换内容</a> 』 『  
<a href="javascript:HighlightAll('form.content')">复制到剪贴板</a> 』</TD></TR>

<TR id=adv style=DISPLAY:none>
<TD vAlign=top align=left width=160 class=a4>


<FONT color=000000><B>投票项目</B><BR>
每行一个投票项目<BR><BR>
<INPUT type=radio CHECKED value=0 name=multiplicity id=multiplicity>
<label for=multiplicity>单选投票</label>
<BR><INPUT type=radio value=1 name=multiplicity id=multiplicity_1> <label for=multiplicity_1>多选投票</label></FONT>

</TD>
<TD width=570 class=a4>
<TEXTAREA onkeydown=presskey(); name=vote rows=5 cols=70><%=Request.form("vote")%></TEXTAREA>
</TD></TR>


<TR>
<TD align=left class=a4>
<IMG src=images/affix.gif><b>增加附件</b></TD>
</TD>

<TD align=left class=a4><font color="FFFFFF"><b><IFRAME src="upfile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></b></font></TD></TR>
<TR>
<TD align=middle class=a3 colSpan=2 height=27><INPUT tabIndex=4 type=submit value=发表新主题 name=submit1> <input name=preview type="submit" value=" 预 览 "> <INPUT type=reset onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" value=" 重 写 "></TD></TR></TBODY></FORM>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<%
htmlend
%>