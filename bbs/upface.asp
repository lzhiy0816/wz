<!-- #include file="setup.asp" -->
<!--#include FILE="inc/upfile.asp"-->

<%
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

if Request("menu")="up" then
id=Conn.Execute("Select id From [user] where username='"&Request.Cookies("username")&"'")(0)

On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set upl = Server.CreateObject("SoftArtisans.FileUp")
If 0 = Err Then
if upl.TotalBytes > int(MaxFace) then error("<li>ͼƬ��С���ó��� "&int(MaxFace/1024)&" KB<br>��ǰ��ͼƬ��СΪ "&int(upl.TotalBytes/1024)&" KB")
if upl.ContentType<>"image/gif" and upl.ContentType<>"image/pjpeg" then error("<li>���ϴ���ͷ����GIF��JPG��ʽ���ļ�")
upl.SaveAs Server.mappath("images/upface/"&id&".gif")
set upl=nothing

else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

set FileUP=new Upload_file 
FileUP.GetDate(-1)
set file=FileUP.file("file")


if file.filesize > int(MaxFace) then error("<li>�ļ���С���ó��� "&int(MaxFace/1024)&" KB<br>��ǰ���ļ���СΪ "&int(file.filesize/1024)&" KB")
if LCase(file.FileExt) <> "gif" and LCase(file.FileExt) <> "jpg" then error("<li>�Բ��𣬱�������ֻ֧��GIF��JPG��ʽ���ļ�<li>��֧�� "&file.FileExt&" ��ʽ���ļ�")
file.SaveToFile Server.mappath("images/upface/"&id&".gif")
set FileUP=nothing

end if

conn.execute("update [user] set userface='images/upface/"&id&".gif' where username='"&Request.Cookies("username")&"'")
message=message&"<li>ͷ���ϴ��ɹ�<li><a href=upface.asp>�����ϴ�ͷ��</a><li><a href=main.asp>������̳��ҳ</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=upface.asp>")

elseif Request("menu")="custom" then
userface=server.htmlencode(Request("userface"))
if userface="" then error("<li>ͷ��URL����û����д")
conn.execute("update [user] set userface='"&userface&"' where username='"&Request.Cookies("username")&"'")


End If



%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>

<title>�ϴ�ͷ��</title>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">��<a href="upface.asp">�ϴ�ͷ��</a></td>
</tr>
</table>

<form enctype=multipart/form-data method=post action=upface.asp?menu=up>
<SCRIPT>valigntop()</SCRIPT>
<table width=97% border=0 cellpadding="0" cellspacing="1" class=a2 height="218" align="center"><tr>
	<td class=a1 height="25" width="544"><p align=center>���ȡ����ġ����������ѡ����Ҫ�ϴ���ͼƬ</p>
	</td><td class=a1><div align=center>��Ŀǰ��ͷ��</span></div></td>
</tr><tr><td class=a3 width="544"><center>
<input type=file name="file" size=33><br><br><input name="Submit" type=submit value=" ȷ �� "></form></td>
		<td rowspan=3 class=a3 align="center"><img src="<%=Conn.Execute("Select userface From [user] where username='"&Request.Cookies("username")&"'")(0)%>"></td>
</tr><tr><td class=a1 height="25" width="544"><center>�Զ���ͷ��</td><tr>
<td class=a3 width="544"><form method=post action=upface.asp?menu=custom><center>ͷ��URL��<input size="35" value="<%=Conn.Execute("Select userface From [user] where username='"&Request.Cookies("username")&"'")(0)%>" name="userface">
<br><br><input name="Submit0" type=submit value=" �� �� "></form></td></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
htmlend
%>