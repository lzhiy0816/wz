<!-- #include file="setup.asp" -->

<%

top

forumid=int(Request("forumid"))
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")


content=server.htmlencode(Request.Form("content"))
content=replace(content,vbCrlf,"<br>")
content=replace(content,"\","&#92;")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" and Request("preview")="" then

If not conn.Execute("Select username From [prison] where username='"&Request.Cookies("username")&"'" ).eof Then error("<li>�����ؽ�<a href=prison.asp>����</a>")

username=Request.Cookies("username")
icon=Request("icon")
topic=server.htmlencode(Trim(Request("topic")))
topic=replace(topic,"\","&#92;")
if topic="" then
message=message&"<li>����û����д"
end if
if content="" then
message=message&"<li>����û����д"
end if

if Len(topic)>50 then
message=message&"<li>�������ⲻ�ܴ��� 50 �ֽ�"
end if

if Len(content)>50000 then
message=message&"<li>�������ݲ��ܴ��� 50000 �ֽ�"
end if

if instr(topic,"	") > 0 then
message=message&"<li>�����в��ܺ��������ַ�"
end if

if Application("content")=content then
message=message&"<li>�벻Ҫ������ͬ������"
end if


''''''''''''''''''''


sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3


if rs("userlife")<5 then
message=message&"<li>��������ֵ < <FONT color=red>5</FONT> ���ܷ�������<li>�����Ե�<A href=shop.asp>�����̵�</A>��������ҩ��<li>ÿ��Чͣ��ʱ��<FONT color=red> 10 </FONT>���ӣ�����ֵ��<FONT color=red>+10</FONT>"
end if

if rs("experience")<1 and Reg10 = 1 then
message=message&"<li>��ע���û���¼����ͣ��<FONT color=red> 10 </FONT>�������ϲſɷ�������"
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
'ͶƱ�������
if Request("vote")<>"" then
vote=server.htmlencode(Trim(Request("vote")))
if instr(vote,"|") > 0 then
message=message&"<li>ͶƱѡ���в��ܺ��С�|���ַ�"
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

message="<li>�����ⷢ��ɹ�<li><a onclick=min_yuzi() target=message href=ShowPost.asp?id="&id&">��������</a><li><a href=ShowForum.asp?forumid="&forumid&">������̳</a><li><a href=main.asp>������̳��ҳ</a>"
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
<title>��������</title>

<CENTER>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<SCRIPT>if ("<%=logo%>"!=''){document.write("<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'")}else{document.write("<img border=0 src=images/lzylogo.gif>")}</SCRIPT>
</td>
<td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/closedfold.gif" border="0">��<a href="ShowForum.asp?forumid=<%=forumid%>"><%=bbsname%></a><br>
����&nbsp;<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">����������</td>
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
<TD vAlign=left class=a1 height=13><b>Ԥ������</b></TD></TR>
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
<TD vAlign=left colSpan=2 height=25 class=a1><b>��������</b></TD></TR>
<TR>
<TD width=160 class=a4 height=25><B>�û�����</B></TD>
<TD width=570 class=a4 height=25><%=Request.Cookies("username")%> [<a href="login.asp?menu=out">�˳���¼</a>]</TD>
</TR>
<TR>
<TD width=160 class=a3 height=25><B>���±��� </B> <SELECT onchange=DoTitle(this.options[this.selectedIndex].value)>
<OPTION value="" selected>&nbsp;����</OPTION> <OPTION
value=[ԭ��]>[ԭ��]</OPTION><OPTION value=[ת��]>[ת��]</OPTION> <OPTION
value=[��ˮ]>[��ˮ]</OPTION><OPTION value=[����]>[����]</OPTION> <OPTION
value=[����]>[����]</OPTION><OPTION value=[�Ƽ�]>[�Ƽ�]</OPTION> <OPTION
value=[����]>[����]</OPTION><OPTION value=[ע��]>[ע��]</OPTION> <OPTION
value=[��ͼ]>[��ͼ]</OPTION><OPTION value=[����]>[����]</OPTION> <OPTION
value=[����]>[����]</OPTION><OPTION value=[����]>[����]</OPTION></SELECT> </TD>
<TD width=570 class=a3 height=25>
<INPUT maxLength=30 size=60 name=topic value="<%=Request("topic")%>"></TD></TR>
<TR>
<TD vAlign=top align=left width=160 class=a4 height=23><B>���ı���</B></TD>
<TD width=570 class=a4>
<script>
index = Math.floor(Math.random() * 24+1);
for(i=1;i<=24;i=i+1) {
if (i==index) ii="checked";else ii=""
document.write("<INPUT type=radio value="+i+" name=icon "+ii+"><IMG src=images/brow/"+i+".gif>��")
if (i==12) document.write("<br>")
}
</script>

 </TD></TR>

<TR>
<TD vAlign=top class=a3 rowSpan=2>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=left border=0>
<TBODY>
<TR>
<TD vAlign=top align=left width=100% class=a3><B>��������</B><BR>
(<a href="javascript:CheckLength();">�鿴���ݳ���</a>)<BR>
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
<INPUT id=advcheck name=advshow type=checkbox value=1 onclick=showadv()><label for=advcheck> ��ʾͶƱѡ��</label>

</TD></TR>

</TBODY></TABLE></TD>
<SCRIPT src="inc/post.js"></SCRIPT>
<SCRIPT src="inc/ybbcode.js"></SCRIPT>
<TD class=a3>
<img border=0 src=images\ybb\bold.gif onclick=YBBbold() alt="����">
<img border=0 src=images\ybb\italicize.gif onclick=YBBitalic() alt="б��">
<img border=0 src=images\ybb\underline.gif onclick=YBBunder() alt="�»���">
<img border=0 src=images\ybb\center.gif onclick=YBBcenter() alt="����">
<img border=0 src=images\ybb\url.gif onclick=YBBurl() alt="���볬����">
<img border=0 src=images\ybb\email.gif onclick=YBBemail() alt="��������ʼ���ַ">
<img border=0 src=images\ybb\image.gif onclick=YBBimage() alt="����ͼƬ">
<img border=0 src=images\ybb\flash.gif onclick=YBBflash() alt="����FLASH�ļ�">
<img border=0 src=images\ybb\rm.gif onclick=YBBrm() alt="����RealPlayer�ļ�">
<img border=0 src=images\ybb\mp.gif onclick=YBBmp() alt="����Media Player�ļ�">
<img border=0 src=images\ybb\quote.gif onclick=YBBquote() alt="����">
<select onchange=ybbfont(this.options[this.selectedIndex].value)>
<option>����</option>
<option value="Arial">Arial</option>
<option value="Arial Black">Arial Black</option>
<option value="Verdana">Verdana</option>
<option value="Times New Roman">Times New Roman</option>
<option value="Garamond">Garamond</option>
<option value="Courier New">Courier New</option>
<option value="Webdings">Webdings</option>
<option value="Wingdings">Wingdings</option>
<option value="����">����</option>
<option value="��Բ">��Բ</option>
<OPTION value="��������">��������</OPTION>
<OPTION value="����Ҧ��">����Ҧ��</OPTION>
<OPTION value="����_GB2312">����_GB2312</OPTION>
<OPTION value="����">����</OPTION>
<OPTION value="���Ĳ���">���Ĳ���</OPTION>
<OPTION value="����ϸ��">����ϸ��</OPTION>
<OPTION value="������κ">������κ</OPTION>
<OPTION value="�����п�">�����п�</OPTION>
<OPTION value="��������">��������</OPTION>
<OPTION value="����_GB2312">����_GB2312</OPTION>
<OPTION value="����">����</OPTION>
<OPTION value="����">����</OPTION>
<OPTION value="������">������</OPTION>
<OPTION value="��Բ">��Բ</OPTION>
</SELECT> <select onchange=ybbsize(this.options[this.selectedIndex].value)>
<OPTION>��С</OPTION>

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

<div id="sending" style="position:absolute;visibility:hidden;"><table width=520 height=170 border=0 cellspacing=2 cellpadding=0 class=a2><tr><td class=a3 align=center>�������ڷ���, ���Ժ�...</td></tr></table></div>

<TEXTAREA onkeydown="presskey();" name=content rows=10 cols=70><%=Request.Form("content")%></TEXTAREA>
<br>
���������������������������� <a href="javascript:replac()">�滻����</a> �� ��  
<a href="javascript:HighlightAll('form.content')">���Ƶ�������</a> ��</TD></TR>

<TR id=adv style=DISPLAY:none>
<TD vAlign=top align=left width=160 class=a4>


<FONT color=000000><B>ͶƱ��Ŀ</B><BR>
ÿ��һ��ͶƱ��Ŀ<BR><BR>
<INPUT type=radio CHECKED value=0 name=multiplicity id=multiplicity>
<label for=multiplicity>��ѡͶƱ</label>
<BR><INPUT type=radio value=1 name=multiplicity id=multiplicity_1> <label for=multiplicity_1>��ѡͶƱ</label></FONT>

</TD>
<TD width=570 class=a4>
<TEXTAREA onkeydown=presskey(); name=vote rows=5 cols=70><%=Request.form("vote")%></TEXTAREA>
</TD></TR>


<TR>
<TD align=left class=a4>
<IMG src=images/affix.gif><b>���Ӹ���</b></TD>
</TD>

<TD align=left class=a4><font color="FFFFFF"><b><IFRAME src="upfile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></b></font></TD></TR>
<TR>
<TD align=middle class=a3 colSpan=2 height=27><INPUT tabIndex=4 type=submit value=���������� name=submit1> <input name=preview type="submit" value=" Ԥ �� "> <INPUT type=reset onclick="{if(confirm('�������Ҫ���ȫ�������ݣ���ȷ��Ҫ�����?')){return true;}return false;}" value=" �� д "></TD></TR></TBODY></FORM>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<%
htmlend
%>