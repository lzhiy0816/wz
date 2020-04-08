<!-- #include file="setup.asp" -->
<!--#include FILE="inc/upfile.asp"-->

<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

if Request("menu")="up" then
id=Conn.Execute("Select id From [user] where username='"&Request.Cookies("username")&"'")(0)

On Error Resume Next
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set upl = Server.CreateObject("SoftArtisans.FileUp")
If 0 = Err Then
if upl.TotalBytes > int(MaxFace) then error("<li>图片大小不得超过 "&int(MaxFace/1024)&" KB<br>当前的图片大小为 "&int(upl.TotalBytes/1024)&" KB")
if upl.ContentType<>"image/gif" and upl.ContentType<>"image/pjpeg" then error("<li>您上传的头像不是GIF、JPG格式的文件")
upl.SaveAs Server.mappath("images/upface/"&id&".gif")
set upl=nothing

else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

set FileUP=new Upload_file 
FileUP.GetDate(-1)
set file=FileUP.file("file")


if file.filesize > int(MaxFace) then error("<li>文件大小不得超过 "&int(MaxFace/1024)&" KB<br>当前的文件大小为 "&int(file.filesize/1024)&" KB")
if LCase(file.FileExt) <> "gif" and LCase(file.FileExt) <> "jpg" then error("<li>对不起，本服务器只支持GIF、JPG格式的文件<li>不支持 "&file.FileExt&" 格式的文件")
file.SaveToFile Server.mappath("images/upface/"&id&".gif")
set FileUP=nothing

end if

conn.execute("update [user] set userface='images/upface/"&id&".gif' where username='"&Request.Cookies("username")&"'")
message=message&"<li>头像上传成功<li><a href=upface.asp>返回上传头像</a><li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=upface.asp>")

elseif Request("menu")="custom" then
userface=server.htmlencode(Request("userface"))
if userface="" then error("<li>头像URL不能没有填写")
conn.execute("update [user] set userface='"&userface&"' where username='"&Request.Cookies("username")&"'")


End If



%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>

<title>上传头像</title>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<a href="upface.asp">上传头像</a></td>
</tr>
</table>

<form enctype=multipart/form-data method=post action=upface.asp?menu=up>
<SCRIPT>valigntop()</SCRIPT>
<table width=97% border=0 cellpadding="0" cellspacing="1" class=a2 height="218" align="center"><tr>
	<td class=a1 height="25" width="544"><p align=center>请点取下面的“浏览”按键选择您要上传的图片</p>
	</td><td class=a1><div align=center>您目前的头像</span></div></td>
</tr><tr><td class=a3 width="544"><center>
<input type=file name="file" size=33><br><br><input name="Submit" type=submit value=" 确 认 "></form></td>
		<td rowspan=3 class=a3 align="center"><img src="<%=Conn.Execute("Select userface From [user] where username='"&Request.Cookies("username")&"'")(0)%>"></td>
</tr><tr><td class=a1 height="25" width="544"><center>自定义头像</td><tr>
<td class=a3 width="544"><form method=post action=upface.asp?menu=custom><center>头像URL：<input size="35" value="<%=Conn.Execute("Select userface From [user] where username='"&Request.Cookies("username")&"'")(0)%>" name="userface">
<br><br><input name="Submit0" type=submit value=" 更 新 "></form></td></table>
<SCRIPT>valignbottom()</SCRIPT>
<%
htmlend
%>