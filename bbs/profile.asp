<!-- #include file="setup.asp" -->
<%
top
username=Request("username")


sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn

if rs.eof then
error("<li>"&username&"的用户资料不存在")
end if


select case rs("sex")
case "male"
sex="男"
case "female"
sex="女"
end select

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title><%=rs("username")%>用户资料</title>
<script src="inc/birth.js"></script>
<script>var badwords= "<%=badwords%>";</script>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　查看<%=rs("username")%>用户资料</td>
</tr>
</table>
<br>
<table width=97% border="0" cellspacing="0" cellpadding="5" align="center">
<tr>
<td><img src="<%=rs("userface")%>" width="32" height="32">　<b><%=rs("username")%></b></td>
<td align="right" valign="bottom"><b>:::相关操作:::</b>　｛
<img src="images/finds.gif"> <a href=searchok.asp?search=author&searchxm=username&content=<%=rs("username")%>> 搜索帖子</a>　
<img src="images/badlist.gif"> <a href=friend.asp?menu=bad&username=<%=rs("username")%>> 加为坏人</a>　

<img src="images/friend1.gif"> <a href=friend.asp?menu=add&username=<%=rs("username")%>> 加为好友</a>　
<img src="images/message1.gif"> <a style=cursor:hand onclick="javascript:open('friend.asp?menu=post&incept=<%=rs("username")%>','','width=320,height=170')"> 发送讯息</a> ｝　</td>
</tr>
</table>

<SCRIPT>valigntop()</SCRIPT>
<table width=97% align=center cellspacing=0 cellpadding=0 border=0>
<tr>
<td class=a2 height="676">
<table width=100% cellspacing=1 cellpadding=6 border=0>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>查看用户“<%=rs("username")%>”的资料</b></td>
</tr>
<tr bgcolor="FFFFFF">
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;<img src="images/1.gif">:::社区信息:::</b></td>
</tr>
<tr>
<td class=a3 height="5" align="left" valign="middle" width="50%"><b>　昵　　称：</b><%=rs("username")%></td>
<td class=a3 height="5" align="left" valign="middle" width="50%"><b>　发表原贴：</b><%=rs("posttopic")%></td>
</tr>
<tr>
<td class=a4 height="2" align="left" valign="middle" width="50%"><b>
　等级名称：</b><Script>document.write(level(<%=rs("experience")%>,<%=rs("membercode")%>,'','')+levelname);</Script></td>
<td class=a4 height="2" align="left" valign="middle" width="50%"><b>　发表回贴：</b><%=rs("postrevert")%></td>
</tr>
<tr class=a4>
<td height="1" align="left" width="50%" class=a3><b>　门　　派：</b><%
if rs("faction")=empty then
response.write "无"
else
response.write rs("faction")
end if
%></td>
<td height="1" align="left" width="50%" class=a3>

<b>　收录精华：</b><%=rs("goodtopic")%></td>
</tr>
<tr class=a4>
<td height="1" align="left" width="50%"><b>　配　　偶：</b><%
if rs("consort")=empty then
response.write "无"
else
response.write "<a href=Profile.asp?username="&rs("consort")&">"&rs("consort")&"</a>"
end if
%></td>
<td height="1" align="left" width="50%">

<b>　被删原贴：</b><%=rs("deltopic")%></td>
</tr>
<tr class=a3>
<td height="2" align="left" width="50%"><b>
　注册日期：</b><%=rs("regtime")%></td>
<td height="2" align="left" width="50%"><b>　社区金币：</b><%=rs("money")%></td>
</tr>
<tr class=a4>
<td height="2" align="left" width="50%"><b>
　上次登录：</b><%=rs("landtime")%></td>
<td height="2" align="left" width="50%"><b>
　体 力 值：</b><%=rs("userlife")%></td>
</tr>
<tr class=a3>
<td height="2" align="left" width="50%"><b>
　登录次数：</b><%=rs("degree")%></td>
<td height="2" align="left" width="50%"><b>　经 验 值：</b><%=rs("experience")%></td>
</tr>

<tr bgcolor="FFFFFF">
<td height="25" align="left" valign="middle" colspan="2"><b>&nbsp;<img src="images/1.gif">:::生活信息:::</b></td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%"> 　<b>真实姓名：</b><%=rs("realname")%>
</td>
<td height="71" align="left" valign="top" class=a3 rowspan="18" width="50%">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td height="153" valign="top" align="center">
<div align=left><b> 　用户照片：</b></div>
<br>

<script>
if("<%=rs("userphoto")%>"!=""){
document.write("<a target=_blank href=<%=rs("userphoto")%>><img src=<%=rs("userphoto")%> onload='javascript:if(this.height>112)this.height=112' border=0></a>")
}
</script>


</td>
</tr>






<tr>
<td height="100">
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>
&nbsp;性　　格：</b><br>
　<table border="0">
<tr>
<td width="100%"><%=rs("character")%></td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td height="120">
<table width="100%" border="0" cellspacing="0" cellpadding="1">
<tr>
<td class=a2 height="100">
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>&nbsp;简　介：</b><br>
<br>
<%=rs("personal")%> </td>
</tr>
</table>
</td>
</tr>
</table>
</td>
</tr>
<tr>
<td height="100">
<table width="100%" border="0" cellspacing="0" cellpadding="1" style=TABLE-LAYOUT:fixed>
<tr>
<td class=a2>
<table width="100%" border="0" cellspacing="0" cellpadding="10">
<tr>
<td class=a4 height="100" valign="top"><b>&nbsp;签名档：</b><br>
<br>
<%=rs("sign")%></td>
</tr>
</table>
</td>
</tr>
</table>


</td>
</tr>


</table>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%"><b>　性　　别：</b><%=sex%> </td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%"><b> 　国　　家：</b><%=rs("country")%>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%">　<b>省　　份：</b><%=rs("province")%> </td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%">　<b>城　　市：</b><%=rs("city")%></td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%">　<b>生　　肖：</b><script>document.write(getpet("<%=rs("birthday")%>"));</script></td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%">　<b>血　　型：</b><%=rs("blood")%> </td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%">　<b>星　　座：</b><script>document.write(astro("<%=rs("birthday")%>"));</script></td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%">　<b>信　　仰：</b><%=rs("belief")%>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%">　<b>职　　业：</b><%=rs("occupation")%>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%">　<b>婚姻状况：</b><%=rs("marital")%>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%">　<b>最高学历：</b><%=rs("education")%>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a4 width="50%">　<b>毕业院校：</b><%=rs("college")%>
</td>
</tr>
<tr class=a4>
<td height="2" valign="top" bgcolor="FFFFFF" width="50%"><b>&nbsp;<img src="images/1.gif">:::联系方法:::</b></td>
</tr>
<tr class=a4>
<td height="1" valign="top" class=a3 width="50%">　<b>ＱＱ号码：</b><a target=_blank href=http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<%=rs("userqq")%>><%=rs("userqq")%></a></td>
</tr>
<tr class=a4>
<td height="1" valign="top" class=a4 width="50%">　<b>ICQ 号码：</b><a target=_blank href=http://web.icq.com/wwp?Uin=<%=rs("icq")%>><%=rs("icq")%></a></td>
</tr>
<tr class=a4>
<td height="2" valign="top" class=a3 width="50%">　<b>电子邮件：</b><a href="mailto:<%=rs("usermail")%>"><%=rs("usermail")%></a></td>
</tr>
<tr class=a4>
<td height="1" valign="top" class=a4 width="50%">　<b>个人主页：</b><a target="_blank" href="<%=rs("userhome")%>"><%=rs("userhome")%></a></td>
</tr>

</table>

</td>
</tr>
</table>


<%rs.close%>

<center>

<script>
if(top!=self){

document.write("<table width=97% cellspacing=1 cellpadding=6 border=0 class=a2><tr><td height=25 align=center class=a1><b>:::最近发表的文章:::</b></td></tr><tr bgcolor=FFFFFF><td>")
<%
sql="select top 5 id,icon,topic,posttime from forum where deltopic<>1 and username='"&username&"' order by id Desc"
rs.Open sql,Conn
Do While Not RS.EOF
%>
document.write("<IMG src=images/brow/<%=rs("icon")%>.gif border=0> <a onclick=min_yuzi() target=message href=ShowPost.asp?id=<%=rs("id")%>><%=rs("topic")%></a> -- <%=rs("posttime")%><br>")
<%
RS.MoveNext
loop
RS.Close   
%>
document.write("</td></tr></table>")
}
</script>
<SCRIPT>valignbottom()</SCRIPT>

<br>
<INPUT onclick=history.back(-1) type=button value=" << 返 回 ">
<%
htmlend
%>