<!-- #include file="setup.asp" -->
<%
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")
id=int(Request("id"))
top



forumid=Conn.Execute("Select forumid From forum where id="&id)(0)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if Request.ServerVariables("request_method") = "POST" then

content=server.htmlencode(Request.Form("content"))
content=replace(content,vbCrlf,"<br>")
content=replace(content,"\","&#92;")
topic=server.htmlencode(Trim(Request("idtopic")))
topic=replace(topic,"\","&#92;")

if content=empty then message=message&"<li>����û����д"

if Len(content)>50000 then message=message&"<li>�������ݲ��ܴ��� 50000 �ֽ�"

if message<>"" then error(""&message&"")

if isnumeric(""&Request("retopicid")&"") then
sql="select * from reforum where id="&Request("retopicid")&""
else
sql="select * from forum where ID="&ID&" and forumid="&forumid&""
end if

rs.Open sql,Conn,1,3
if rs("username")<>Request.Cookies("username") and instr(Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0),"|"&Request.Cookies("username")&"|")=0 and membercode<4 then
error("<li>�Բ�������Ȩ�޲�����")
end if


if topic<>empty then rs("topic")=""&topic&""

content=""&content&"<br><br>�۴������ѱ� "&Request.Cookies("username")&" �� "&now()&" �༭����"
rs("content")=content
rs.update
rs.close

if Request.Cookies("filename")<>empty then
master=split(Request.Cookies("filename"),"|")
for i = 1 to ubound(master)
conn.execute("update [upfile] set id='"&id&"' where filename='"&master(i)&"'")
next
Response.Cookies("filename")=""
end if
 
message="<li>�޸����ӳɹ�<li><a onclick=min_yuzi() target=message href=ShowPost.asp?id="&id&">��������</a><li><a href=ShowForum.asp?forumid="&forumid&">������̳</a><li><a href=main.asp>����������ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=ShowForum.asp?forumid="&forumid&">")

end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sql="select * from bbsconfig where id="&forumid&""
rs.Open sql,Conn
bbsname=rs("bbsname")
logo=rs("logo")
rs.close

if isnumeric(Request("retopicid")) then
sql="select * from reforum where id="&Request("retopicid")&""
else
sql="select * from forum where ID="&id&" and forumid="&forumid&""
end if
rs.Open sql,Conn
if rs.eof then
error("<li>���ݿ��в����ڴ����ӵ�����")
end if

if rs("username")<>Request.Cookies("username") and instr(Conn.Execute("Select moderated From [bbsconfig] where id="&forumid&" ")(0),"|"&Request.Cookies("username")&"|")=0 and membercode<4 then
error("<li>�Բ�������Ȩ�޲�����")
end if

content =rs("content")
rs.close

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>


<script>function runEx(){var winEx = window.open('', '', 'resizable,scrollbars');winEx.document.write(unescape(event.srcElement.parentElement.children[0].value));}</script>


<title>�༭����</title>

<CENTER>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%">
<SCRIPT>if ("<%=logo%>"!=''){document.write("<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'")}else{document.write("<img border=0 src=images/lzylogo.gif>")}</SCRIPT>
</td>
<td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/closedfold.gif" border="0">��<a href="ShowForum.asp?forumid=<%=forumid%>"><%=bbsname%></a><br>

���� <img src=images/bar.gif><img src=images/openfold.gif>��<a onclick=min_yuzi() target=message href=ShowPost.asp?id=<%=id%>><%=Request("topic")%></a>
</td>
</tr>
</table>
<br>

<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=5 width=97% border=0 class=a2>
<form method=post name=form onsubmit="return ValidateForm()">
<TBODY>
<TR>
<TD id=titlelarge vAlign=left colSpan=2 height=25 class=a1><b>�༭����</b></TD></TR>

<%if not isnumeric(Request("retopicid")) then%>

<TR class=a4>
<TD width=160 height=25><B>���±���</B>&nbsp; </TD>
<TD width=574 height=25>
<INPUT maxLength=30 size=60 name=idtopic value="<%=Request("topic")%>"></TD></TR>


<%end if%>

<TR>
<TD vAlign=top class=a3 rowSpan=2>
<TABLE cellSpacing=0 cellPadding=0 width=100% align=left border=0>
<TBODY>
<TR>
<TD vAlign=top align=left width=100% class=a3><B>��������</B><BR>
(<a href="javascript:CheckLength();">�鿴���ݳ���</a>)<BR><BR>
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
</SELECT>
<select onchange=ybbsize(this.options[this.selectedIndex].value)>
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
	<IMG height=10 width=20 border=1></TD></TD><SCRIPT>rgb(65,1,10)</SCRIPT></TR>
	</TBODY>
</TABLE>
</TD></TR>
<TR>
<TD width=574 class=a3>
<div id="sending" style="position:absolute;visibility:hidden;"><table width=520 height=170 border=0 cellspacing=2 cellpadding=0 class=a2><tr><td class=a3 align=center>�������ڷ���, ���Ժ�...</td></tr></table></div>

<TEXTAREA onkeydown=presskey(); name=content rows=10 cols=70><%=content%></TEXTAREA>
<br>���������������������������� <a href="javascript:replac()">�滻����</a> �� ��  
<a href="javascript:HighlightAll('form.content')">���Ƶ�������</a> ��</TD></TR>

<TR>
<TD align=left class=a4>
<IMG src=images/affix.gif><b>���Ӹ���</b></TD>

<TD align=left class=a4><font color="FFFFFF"><b><IFRAME src="upfile.asp" frameBorder=0 width="100%" scrolling=no height=21></IFRAME></b></font></TD></TR>


<TR>
<TD align=middle class=a3 colSpan=2 height=27><INPUT tabIndex=4 type=submit value=ȷ���༭ name=submit1>&nbsp; <INPUT type=reset value=" �� д "></TD></TR></TBODY></FORM>
</TABLE>
<SCRIPT>valignbottom()</SCRIPT>

<Script>
document.form.content.value = unybbcode(document.form.content.value);
document.form.content.focus();
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script>

<%
htmlend
%>