<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->

<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")



select case Request("menu")

case "editProfileok"
editProfileok

case "contactok"
contactok

case "passok"
passok

end select


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<SCRIPT src="inc/birthday.js"></SCRIPT>
<title>控制面板</title>
<title><%=bbsname%></title>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif">　控制面板</td>
</tr>
</table>
<br>


<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0" class="a2">
  <TR id=TableTitleLink class=a1 height="25">
      <Td width="14%" align="center"><b><a href="usercp.asp">控制面板首页</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp">基本资料修改</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp?menu=contact">编辑论坛选项</a></b></td>
      <TD width="14%" align="center"><b><a href="EditProfile.asp?menu=pass">用户密码修改</a></b></td>
      <TD width="14%" align="center"><b><a href="message.asp">用户短信服务</a></b></td>
      <TD width="14%" align="center"><b><a href="friend.asp">编辑好友列表</a></b></td>
      <TD width="16%" align="center"><b><a href="favorites.asp">用户收藏管理</a></b></td></TR></TABLE>
<br>
<%


select case Request("menu")
case ""
index
case "contact"
contact
case "pass"
pass
end select




sub index
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
%>
<script>function New(para_URL){var URL=new String(para_URL);window.open(URL,'','resizable,scrollbars')}</script>


<SCRIPT>valigntop()</SCRIPT>

<table width=97% cellspacing=1 cellpadding=4 align=center border=0 class=a2>
  <form method="POST" name="form" action="EditProfile.asp?menu=editProfileok">

<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>
基本资料</b></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="30%"> 　<b>真实姓名：</b> 
      <input type="text" name="realname" size="17" value="<%=rs("realname")%>">
</td>
    <td height="71" align="left" valign="top" class=a4 rowspan="11" width="70%"> 
      <table width="486" cellspacing="0" cellpadding="5">
<tr>
<td width="476"><b>头<font color="000000">　</font>像：</b>
<img src="<%=rs("userface")%>" name="tus" width="32" height="32">　<script>function showimage(){document.images.tus.src=""+document.form.userface.options[document.form.userface.selectedIndex].value+"";}</script><select name=userface size=1 onChange="showimage()">
<option value="<%=rs("userface")%>">默认</option>
<script>
for(i=1;i<=84;i=i+1) {
document.write("<option value='images/face/"+i+".gif'>"+i+"</option>")
}
</script>
</select>　<a href="JavaScript:New('face.htm')">查看所有的头像列表</a></td>
</tr>
<tr>
<td width="466"><b>生<font color="000000">　</font>日：</b> <input onfocus="show_cele_date(birthday,'','',birthday)" value="<%=rs("birthday")%>" name="birthday"></td>
</tr>
<tr>
<td width="476"><b><font color="000000">性　格</font>： &nbsp; </b>

<textarea name=character rows=5 cols=65><%=rs("character")%></textarea><font color="000000">　 </font>
</td>
</tr>
<tr>
<td width="476"><b>个人简介： &nbsp;
<textarea name=personal rows=5 cols=65><%=rs("personal")%></textarea>
</b></td>
</tr>
<tr>
<td height="10" width="476"><b>签名档： &nbsp;
<textarea name=sign rows=5 cols=65><%=rs("sign")%></textarea> </b> </td>
</tr>
</table>
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="30%"><b> 　性　　别：</b> 
      <select size=1 name=sex>
<option value="" selected>[请选择]</option>
<option value=male <%if rs("sex")="male" then%>selected<%end if%>>男</option>
<option value=female <%if rs("sex")="female" then%>selected<%end if%>>女</option>
</select>



</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="30%">　<b>国　　家：</b> <b> 
      <input type="text" name="country" size="17" value="<%=rs("country")%>">
</b> </td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="30%">　<b>省　　份：</b> 
      <input type="text" name="province" size="17" value="<%=rs("province")%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="30%">　<b>城　　市： </b> 
      <input type="text" name="city" size="17" value="<%=rs("city")%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="30%">　<b>血　　型：</b> 
      <input maxlength=4 size=4
name=blood value="<%=rs("blood")%>">
</td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="30%">　<b>信　　仰：</b> 
      <input type="text" name="belief" size="17" value="<%=rs("belief")%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="30%">　<b>职　　业：</b> 
      <input type="text" name="occupation" size="17" value="<%=rs("occupation")%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="30%">　<b>婚姻状况：</b> 
      <input maxlength=4 size=4
name=marital value="<%=rs("marital")%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a4 width="30%">　<b>最高学历：</b> 
      <input type="text" name="education" size="17" value="<%=rs("education")%>"></td>
</tr>
<tr>
    <td height="2" valign="top" class=a3 width="30%">　<b>毕业院校：</b> 
      <input type="text" name="college" size="17" value="<%=rs("college")%>"></td>
</tr>
  <form method="POST" name="form"action="EditProfile.asp">
<input type=hidden name="menu" value="contactok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>联系资料</b></td>
</tr>
<tr>
    <td height="2" align="right" width="45%" class=a4><b>QQ号码：</b></td>
    <td height="2" align="left" width="55%" class=a4> &nbsp; 
      <input type="text" name="userqq" size="40" value="<%=rs("userqq")%>">
</td>
</tr>
<tr>
    <td height="2" align="right" width="45%" class=a3><b>ICQ号码：</b></td>
    <td height="2" align="left" width="55%" class=a3> &nbsp; 
      <input type="text" name="icq" size="40" value="<%=rs("icq")%>">
</td>
</tr>
<tr>
    <td class=a4 height="2" align="right" valign="middle" width="45%"><b>Email地址：</b></td>
    <td class=a4 height="2" align="left" valign="middle" width="55%"> &nbsp; 
      <input type="text" name="usermail" size="40" value="<%=rs("usermail")%>">
</td>
</tr>

<tr class=a3>
    <td height="1" align="right" width="45%"><b>个人主页：</b></td>
    <td height="1" align="left" width="55%"> &nbsp; 
      <input type="text" name="userhome" size="40" value="<%=rs("userhome")%>">
</td>
</tr>

<tr class=a4>
    <td height="1" align="center" width="100%" colspan="2">

<input type="submit" name="Submit1" value=" 确 定 "></td>
</tr>
</table>

<SCRIPT>valignbottom()</SCRIPT>

</form>


&nbsp;<Script>
document.form.sign.value = unybbcode(document.form.sign.value);
document.form.sign.focus();
function unybbcode(temp) {
temp = temp.replace(/<br>/ig,"\n");
return (temp);
}
</Script><%
end sub





sub editProfileok
username=Request.Cookies("username")
realname=server.htmlencode(Request("realname"))
birthday=server.htmlencode(Request("birthday"))
userface=server.htmlencode(Request("userface"))
sex=server.htmlencode(Request("sex"))
country=server.htmlencode(Request("country"))
province=server.htmlencode(Request("province"))
city=server.htmlencode(Request("city"))
blood=server.htmlencode(Request("blood"))
character=server.htmlencode(Request("character"))
belief=server.htmlencode(Request("belief"))
college=server.htmlencode(Request("college"))
marital=server.htmlencode(Request("marital"))
education=server.htmlencode(Request("education"))
personal=server.htmlencode(Request("personal"))
occupation=server.htmlencode(Request("occupation"))
usermail=server.htmlencode(Request("usermail"))
userhome=server.htmlencode(Request("userhome"))
userqq=server.htmlencode(Request("userqq"))
icq=server.htmlencode(Request("icq"))
sign=server.htmlencode(Request("sign"))
sign=replace(sign,vbCrlf,"<br>")
sign=replace(sign,"\","&#92;")


temp=UCase(sign)
if instr(temp,"[FLASH]")>0 or instr(temp,"[RM]")>0 or instr(temp,"[MP]")>0 then
message=message&"<li>签名档中不能含有[FLASH] [RM] [MP]代码"
end if

if Len(character)>255 then
message=message&"<li>性格不能大于 255 个字符"
end if

if Len(personal)>255 then
message=message&"<li>个人简介不能大于 255 个字符"
end if

if Len(sign)>255 then
message=message&"<li>签名档不能大于 255 个字节"
end if

if not isnumeric(userqq) and userqq<>"" then
message=message&"<li>QQ号码请用数字填写"
end if

if not isnumeric(icq) and icq<>"" then
message=message&"<li>ICQ号码请用数字填写"
end if

if instr(usermail,"@")=0 then
message=message&"<li>您的电子邮件地址填写错误"
end if


if message<>"" then
error(""&message&"")
end if

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3
rs("realname")=realname
rs("birthday")=birthday
rs("userface")=userface
rs("sex")=sex
rs("country")=country
rs("province")=province
rs("city")=city
rs("blood")=blood
rs("character")=character
rs("belief")=belief
rs("college")=college
rs("marital")=marital
rs("education")=education
rs("personal")=personal
rs("sign")=sign
rs("occupation")=occupation
rs("usermail")=usermail
rs("userhome")=userhome
if icq<>empty then rs("icq")=icq
if userqq<>empty then rs("userqq")=userqq
rs("landtime")=now()
rs.update
rs.close
message=message&"<li>修改资料成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub





sub contact
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
%>

<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="EditProfile.asp">
<input type=hidden name="menu" value="contactok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>编辑论坛选项</b></td>
</tr>
<tr>
    <td height="2" align="right" width="45%" class=a4><b>YBB代码：</b></td>
    <td height="2" align="left" width="55%" class=a4> &nbsp; 
      
<input type=radio name="ybbcode" value="0" <%if Request.Cookies("ybbcode")="0" then%>checked<%end if%> id=ybbcode><label for=ybbcode>关闭</label>
<input type=radio name="ybbcode" value="1" <%if Request.Cookies("ybbcode")="1" then%>checked<%end if%> id=ybbcode_1><label for=ybbcode_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
      
</td>
</tr>
<tr>
    <td height="2" align="right" width="45%" class=a3><b>[IMG]代码：</b></td>
    <td height="2" align="left" width="55%" class=a3> &nbsp; 
<input type=radio name="ybbimg" value="0" <%if Request.Cookies("ybbimg")="0" then%>checked<%end if%> id=ybbimg><label for=ybbimg>关闭</label>
<input type=radio name="ybbimg" value="1" <%if Request.Cookies("ybbimg")="1" then%>checked<%end if%> id=ybbimg_1><label for=ybbimg_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
</td>
</tr>
<tr>
    <td class=a4 height="2" align="right" valign="middle" width="45%"><b>
    [FLASH]代码：</b></td>
    <td class=a4 height="2" align="left" valign="middle" width="55%"> &nbsp; 
<input type=radio name="ybbflash" value="0" <%if Request.Cookies("ybbflash")="0" then%>checked<%end if%> id=ybbflash><label for=ybbflash>关闭</label>
<input type=radio name="ybbflash" value="1" <%if Request.Cookies("ybbflash")="1" then%>checked<%end if%> id=ybbflash_1><label for=ybbflash_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
</td>
</tr>

<tr class=a3>
    <td height="2" align="right" width="45%"><b>表情代码：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="ybbbrow" value="0" <%if Request.Cookies("ybbbrow")="0" then%>checked<%end if%> id=ybbbrow><label for=ybbbrow>关闭</label>
<input type=radio name="ybbbrow" value="1" <%if Request.Cookies("ybbbrow")="1" then%>checked<%end if%> id=ybbbrow_1><label for=ybbbrow_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font>
</td>
</tr>


<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>编辑帖子选项</b></td>
</tr>


<tr class=a4>
    <td height="2" align="right" width="45%"><b>在帖子中显示用户头像：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="showface" value="0" <%if Request.Cookies("showface")="0" then%>checked<%end if%> id=showface><label for=showface>关闭</label>
<input type=radio name="showface" value="1" <%if Request.Cookies("showface")="1" then%>checked<%end if%> id=showface_1><label for=showface_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font></td>
</tr>


<tr class=a3>
    <td height="2" align="right" width="45%"><b>在帖子中显示用户签名：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="sign" value="0" <%if Request.Cookies("sign")="0" then%>checked<%end if%> id=sign><label for=sign>关闭</label>
<input type=radio name="sign" value="1" <%if Request.Cookies("sign")="1" then%>checked<%end if%> id=sign_1><label for=sign_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[打开]</font></td>
</tr>

<tr class=a4>
    <td height="2" align="right" width="45%"><b>总是使用框架显示帖子：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="frame" value="0" <%if Request.Cookies("frame")="0" then%>checked<%end if%> id=frame><label for=frame>关闭</label>
<input type=radio name="frame" value="1" <%if Request.Cookies("frame")="1" then%>checked<%end if%> id=frame_1><label for=frame_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[关闭]</font></td>
</tr>

<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>编辑个性选项</b></td>
</tr>

<tr class=a3>
    <td height="2" align="right" width="45%"><b>每页显示主题数：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<select name="pagesetup" size="1">
<option value="">默认</option>
<option value="10" <%if Request.Cookies("pagesetup")="10" then%>selected<%end if%>>10</option>
<option value="20" <%if Request.Cookies("pagesetup")="20" then%>selected<%end if%>>20</option>
<option value="30" <%if Request.Cookies("pagesetup")="30" then%>selected<%end if%>>30</option>
</select>
    </td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b>主题列表中突出关键词：</b><br>设定的关键词将用<font color="FF0000">红色</font>显示&nbsp;&nbsp; </td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input size=20 name=key_word value="<%=Request.Cookies("key_word")%>"></td>
</tr>


<tr class=a3>
    <td height="2" align="right" width="45%"><b>过滤用户名单：</b><br><font color="FF0000">多人请用“|”分隔</font>&nbsp;&nbsp; </td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input size=30 name=badlist value="<%=Request.Cookies("badlist")%>"></td>
</tr>




<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>个人状态设定</b></td>
</tr>


<tr class=a4>
    <td height="2" align="right" width="45%"> <b>在线状态：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
<input type=radio name="eremite" value="0" <%if Request.Cookies("eremite")="0" then%>checked<%end if%> id=eremite><label for=eremite>上线</label>
<input type=radio name="eremite" value="1" <%if Request.Cookies("eremite")="1" then%>checked<%end if%> id=eremite_1><label for=eremite_1>隐身</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[上线]</font></td>
</tr>

<tr class=a3>
    <td height="1" align="right" width="45%"><b>收到短讯时弹出提示框：</b></td>
    <td height="1" align="left" width="55%"> &nbsp; 
<input type=radio name="newmessage" value="0" <%if Request.Cookies("newmessage")="0" then%>checked<%end if%> id=newmessage><label for=newmessage>关闭</label>
<input type=radio name="newmessage" value="1" <%if Request.Cookies("newmessage")="1" then%>checked<%end if%> id=newmessage_1><label for=newmessage_1>打开</label> 
    <font color="FF0000">&nbsp;&nbsp; 默认:[关闭]</font></td>
</tr>


<tr class=a4>
    <td height="1" align="center" width="100%" colspan="2">
<input type="submit" name="Submit2" value=" 确 定 "></td>
</tr>


</table>
<SCRIPT>valignbottom()</SCRIPT>
<center><br>
&nbsp;</form><%
end sub



sub contactok


Response.Cookies("ybbcode")=Request("ybbcode")
Response.Cookies("ybbimg")=Request("ybbimg")
Response.Cookies("ybbflash")=Request("ybbflash")
Response.Cookies("ybbbrow")=Request("ybbbrow")
Response.Cookies("eremite")=Request("eremite")
Response.Cookies("newmessage")=Request("newmessage")
Response.Cookies("showface")=Request("showface")
Response.Cookies("sign")=Request("sign")
Response.Cookies("frame")=Request("frame")
Response.Cookies("badlist")=Request("badlist")
Response.Cookies("key_word")=Request("key_word")
Response.Cookies("pagesetup")=Request("pagesetup")



Response.Cookies("ybbcode").Expires=date+365
Response.Cookies("ybbimg").Expires=date+365
Response.Cookies("ybbflash").Expires=date+365
Response.Cookies("ybbbrow").Expires=date+365
Response.Cookies("eremite").Expires=date+365
Response.Cookies("newmessage").Expires=date+365
Response.Cookies("showface").Expires=date+365
Response.Cookies("sign").Expires=date+365
Response.Cookies("frame").Expires=date+365
Response.Cookies("badlist").Expires=date+365
Response.Cookies("key_word").Expires=date+365
Response.Cookies("pagesetup").Expires=date+365
message=message&"<li>更新设置成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub



sub pass
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn
%>
<SCRIPT>valigntop()</SCRIPT>
<table width=97% cellspacing=1 cellpadding=4 border=0 class=a2 align=center>
  <form method="POST" name="form"action="EditProfile.asp">
<input type=hidden name="menu" value="passok">
<tr>
<td height="20" align="center" colspan="2" valign="bottom" class=a1>
<b>用户密码修改</b></td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b> 原密码：</b></td>
    <td height="2" align="left" width="55%"> &nbsp; 
      <input type="password" name="olduserpass" size="40">
</td>
</tr>
<tr class=a3>
    <td height="2" align="right" width="45%"><b> 新密码：</b><br>
      <font color="#FF0000">如不更改密码此处请留空</font></td>
    <td height="2" align="left" width="55%"> &nbsp; 
      <input type="password" name="userpass" size="40" maxlength="16">
</td>
</tr>
<tr class=a4>
    <td height="2" align="right" width="45%"><b>确认新密码：</b><br>
      <font color="#FF0000">请与您的新密码保持一致</font></td>
    <td height="2" align="left" valign="middle" width="55%"> &nbsp; 
      <input type="password" name="userpass2" size="40" maxlength="16"></td>
</tr>
<tr class=a3>
    <td height="1" align="right" width="45%"><b>密码提示问题：</b></td>
    <td height="1" align="left" width="55%"> &nbsp; 
      <input type="text" name="question" size="40" value="<%=rs("question")%>">
</td>
</tr>
<tr class=a4>
    <td height="1" align="right" width="45%"><b>密码提示答案：</b><br>
	<font color="#FF0000">如不更改答案此处请留空</font></td>
    <td height="1" align="left" width="55%" class=a4> &nbsp; 
      <input type="text" name="answer" size="40" value=""></td>
</tr>
<tr class=a4>
    <td align="center" width="100%" colspan="2">
<input type="submit" name="Submit" value=" 确 定 "></td>
</tr>
</table>
<SCRIPT>valignbottom()</SCRIPT>
<center><br>
&nbsp;</form>
<%
end sub

sub passok
username=Request.Cookies("username")
userpass=Trim(Request("userpass"))
olduserpass=Trim(Request("olduserpass"))
userpass2=Trim(Request("userpass2"))
question=server.htmlencode(Request("question"))

sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn,1,3


if md5(olduserpass)<>rs("userpass") then
message=message&"<li>您的原密码错误"
end if

if userpass<>userpass2 then
message=message&"<li>您的新密码和确认新密码不同"
end if

if message<>"" then
error(""&message&"")
end if

rs("question")=question
if userpass<>empty then rs("userpass")=md5(userpass)
if Request("answer")<>empty then rs("answer")=md5(Request("answer"))

rs("landtime")=now()
rs.update
rs.close
message=message&"<li>修改资料成功<li><a href=usercp.asp>控制面板首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=usercp.asp>")
end sub


htmlend
%>