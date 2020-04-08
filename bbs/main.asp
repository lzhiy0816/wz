<!-- #include file="setup.asp" -->
<%
id=int(Request("id"))
if ""&adminpassword&""=empty then response.redirect "install.asp"
top%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title><%=clubname%> - Powered By BBSxp</title>
<table width="97%" align="center" border="0"><tr><td valign="top"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td><td align="right">
<a href="ShowBBS.asp?menu=2">本周热门帖子</a> | <a href="ShowBBS.asp?menu=1">本周人气帖子</a>
<br><a href="ShowBBS.asp">社区新帖</a> | <a href="ShowBBS.asp?menu=3">精华帖子</a> | <a href="ShowBBS.asp?menu=4">投票帖子</a>
<%
response.write "<br>注册会员:<b>"&conn.execute("Select count(id)from [user]")(0)&"</b>"
if Application("NewUserName")<>empty then response.write " | 新会员:<b><a href=Profile.asp?username="&Application("NewUserName")&">"&Application("NewUserName")&"</a></b>"
%>
</td></tr></table><table height="28" cellspacing="1" cellpadding="1" width="97%" align="center" border="0"><tr><td align="middle" width="3%"><img src="images/announce.gif" align="middle"></td><td width="80%"><marquee onmouseover="this.stop()" onmouseout="this.start()" width="80%" scrollamount="3"><a style=cursor:hand onclick="javascript:open('affiche.asp','','width=400,height=180,resizable,scrollbars')"><%=affichetitle%></a>　［<%=affichetime%>］</marquee></td>
	<td align="right">
	<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}" size=1><option>自选风格</option><option value="cookies.asp?menu=skins">默认风格</option><option value="cookies.asp?menu=skins&no=1">橘色风格</option><option value="cookies.asp?menu=skins&no=2">绿色风格</option><option value="cookies.asp?menu=skins&no=3">灰色风格</option></select> </td></tr></table>
<SCRIPT>valigntop()</SCRIPT>

<table cellspacing="1" cellpadding="4" width="97%" align="center" border="0" class="a2"><tr class="a1" height=25><td valign="center" align="middle" width="5%">状态</td><td valign="center" align="middle" width="47%">论坛</td><td valign="center" align="middle" width="11%">版主</td><td valign="center" align="middle" width="7%">主题</td><td valign="center" align="middle" width="7%">帖子</td><td valign="center" align="middle" width="17%">最后更新</td><td valign="center" align="middle" width="7%">珍藏</td></tr><%
if Request.Cookies("forumid")<>empty then
sql="select * from bbsconfig where id="&Request.Cookies("forumid")&" and hide=1"
rs1.Open sql,Conn
if not rs1.eof then bbslist
rs1.close
end if


  
rs.Open "class order by id",Conn
do while not rs.eof


%> <tr><td colspan="7" height="25" bgcolor="FFFFFF"><b>　<a href="?id=<%=rs("id")%>"><%if id=""&rs("id")&"" then%><img src="images/2.gif" border="0"><%else%><img src="images/1.gif" border="0"><%end if%><%=rs("classname")%></a></b></td></tr><%
		
if rs("id")=id or allclass="1" then
sql="select * from bbsconfig where classid="&rs("id")&" and hide=0"
rs1.Open sql,Conn
do while not rs1.eof

if Not Response.IsClientConnected then responseend

bbslist


rs1.movenext
loop
rs1.close
end if

rs.movenext
loop
rs.close
%> </table>

<SCRIPT>valignbottom()</SCRIPT>
<br>
<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0" class="a2"><tr><td height="25" class="a1" colspan="2">　■<b> </b>您的个人状态</td></tr><tr>
		<td align="middle" width="5%" bgcolor=FFFFFF><img src="images/mystate.gif"></td><td class="a4"><table cellspacing="0" cellpadding="3" width="100%" border="0" style="border-left: 0px none; border-top: 0px none; border-bottom: 1px none"><tr><td><table border="0" cellpadding="0" cellspacing="0" width="100%"><%
if Request.Cookies("username") <> "" then

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn


%> <tr><td width="25%"><%=rs("username")%> - 有 <font face="Georgia, Times New Roman, Times, serif"><%=rs("newmessage")%> 条新留言</font></td><td width="20%">金币：<font face="Georgia, Times New Roman, Times, serif"><%=rs("money")%></font></td><td width="20%">体力：<font face="Georgia, Times New Roman, Times, serif"><%=rs("userlife")%></font></td><td width="17%">发表文章：<font face="Georgia, Times New Roman, Times, serif"><%=rs("posttopic")%></font></td><td width="18%">被删文章：<font face="Georgia, Times New Roman, Times, serif"><%=rs("deltopic")%></font></td></tr><tr><td>等级名称：<Script>document.write(level(<%=rs("experience")%>,<%=rs("membercode")%>,'','')+levelname);</Script></td><td>存款：<font face="Georgia, Times New Roman, Times, serif"><%=rs("savemoney")%></font></td><td>经验：<font face="Georgia, Times New Roman, Times, serif"><%=rs("experience")%></font></td><td>回复文章：<font face="Georgia, Times New Roman, Times, serif"><%=rs("postrevert")%></font></td><td>精华文章：<font face="Georgia, Times New Roman, Times, serif"><%=rs("goodtopic")%> </font></td></tr><%
rs.close
end if
acturl="main.asp"
bbsname=clubname
forumid=0

%> <!-- #include file="inc/line.asp" --><tr><td>
<script>
temp=navigator.appVersion
temp=temp.replace(")","");
temp=temp.replace("MSIE","Internet Explorer");
temp=temp.replace("NT 5.0","2000");
temp=temp.replace("NT 5.1","XP");
temp=temp.replace("NT 5.2","2003");
var appVersion= temp.split ('; '); 
</script>
您的ＩＰ：<font face="Georgia, Times New Roman, Times, serif"><%=remoteaddr%></font></td><td>端口：<font face="Georgia, Times New Roman, Times, serif"><%=Request.ServerVariables("REMOTE_PORT")%></font></td><td>操作系统：<script>document.write(appVersion[2])</script></td><td colspan="2">浏览器：<script>document.write(appVersion[1])</script></td></tr></table></td></tr></table></td></tr></table><%
regonline=conn.execute("Select count(sessionid)from online where username<>''")(0)

if Application("BestOnline")<onlinemany then
Application("BestOnline")=onlinemany
Application("BestOnlineTime")=now()
end if

%><SCRIPT>valignbottom()</SCRIPT>
<br>
<SCRIPT>valigntop()</SCRIPT><table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr><td height="25" class="a1" colspan="2">　<b>■ </b>在线统计</td></tr><tr><td align="middle" width="5%" bgcolor=FFFFFF><img src="images/online.gif"></td><td valign="top" class="a4"><table cellspacing="0" cellpadding="3" width="100%"><tr><td height="15">&nbsp;<img loaded="no" src="images/plus.gif" id="followImg0" style="cursor:hand;" onclick="loadThreadFollow(0,<%=forumid%>)"> 目前论坛总共有 <b><%=onlinemany%></b> 人在线。其中注册用户 <b><%=regonline%></b> 人，访客 <b><%=onlinemany-regonline%></b> 
	人。近日最高在线纪录 <font color=red><b><%=Application("BestOnline")%></b></font> 人，发生在 <%=Application("BestOnlineTime")%>
</td></tr><tr height="25" style="display:none" id="follow0"><td id="followTd0" align="left" class="a4" width="94%" colspan="5"><div onclick="loadThreadFollow(0,<%=forumid%>)"><table width="100%" cellpadding="10"><tr><td width="100%">Loading...</td></tr></table></div></td></tr></table></td></tr></table>

<SCRIPT>valignbottom()</SCRIPT>
<br>
<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="5" width="97%" border="0" class="a2" align="center"><tr><td height="25" colspan="2" class="a1">　■<b> </b>友情链接</td></tr><tr>
<td align="center" bgcolor=FFFFFF width="5%" rowspan="2"><img src="images/shareforum.gif"></td>
<td class="a4"><%
rs.Open "link",Conn
do while not rs.eof
if rs("logo")="" or rs("logo")="http://" then
link1=link1+"<a title='"&rs("intro")&"' href="&rs("url")&" target=_blank>"&rs("name")&"</a> "
else
link2=link2+"<a title='"&rs("name")&""&chr(10)&""&rs("intro")&"' href="&rs("url")&" target=_blank><img src="&rs("logo")&" border=0 width=88 height=31></a> "
end if
rs.movenext
loop
rs.close
%>
<%=link1%>
</td></tr><tr>
<td class="a4"><%=link2%></td></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
<br>
<center>&nbsp;<img src="images/pass0.gif" alt="禁止浏览帖子">&nbsp;关闭论坛&nbsp;&nbsp;<img src="images/pass1.gif" alt="开放浏览帖子">&nbsp;正规论坛&nbsp;&nbsp;<img src="images/pass2.gif" alt="只有登陆用户才能浏览帖子">&nbsp;会员论坛&nbsp;&nbsp;<img src="images/pass3.gif" alt="嘉宾以上的等级才能浏览帖子">&nbsp;嘉宾论坛

<script>
function loadThreadFollow(ino,online){
	var targetImg =eval("followImg" + ino);
	var targetDiv =eval("follow" + ino);
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			targetImg.src="images/minus.gif";
			if (targetImg.loaded=="no"){document.frames["hiddenframe"].location.replace("loading.asp?id="+ino+"&forumid="+online+"");}
		}else{
			targetDiv.style.display="none";
			targetImg.src="images/plus.gif";
		}
}
</script>
<iframe height="0" width="0" name="hiddenframe"></iframe><%
htmlend
sub bbslist
%><tr><td width="5%" height="52" align="middle"bgcolor=FFFFFF>
<img src=images/pass<%=rs1("pass")%>.gif>
</td><td width="47%" valign="top" class="a4"><table width="100%" height="0%" border="0" cellpadding="0" cellspacing="0"><tr><td width="80%" height="50%">『 <a href="ShowForum.asp?forumid=<%=rs1("id")%>"><%=rs1("bbsname")%></a> 』</td><td align="center" width="20%"><a href="ShowForum.asp?forumid=<%=rs1("id")%>&TimeLimit=1"><img alt="查看该论坛当天的帖子" src="images/today.gif" border="0"></a> <a href="ShowForum.asp?forumid=<%=rs1("id")%>&search=goodtopic"><img alt="查看该论坛的精品文章" src="images/goodtopic.gif" border="0"></a> <a href="newtopic.asp?forumid=<%=rs1("id")%>"><img alt="在该论坛发表文章" src="images/postdirect.gif" border="0"></a></font></td></tr><tr><td height="50%" colspan="2">
		<table border="0" width="100%" cellspacing="0">
			<tr>
<SCRIPT>if ("<%=rs1("icon")%>"!=''){
document.write("<td><img src=<%=rs1("icon")%> onload='javascript:if(this.width>200)this.width=200;if(this.height>60)this.height=60;'></td>")}
</SCRIPT>
<td valign="center" width=100%>
<%=rs1("intro")%>
</font></td>
			</tr>
		</table>
		</td></tr></table></td><td valign="center" align="middle" width="11%" class="a3">

<SCRIPT>
var moderated="<%=rs1("moderated")%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i=i+1) {
if (list[i] !=""){document.write("<a href=profile.asp?username="+list[i]+">"+list[i]+"</a><br>")}
}
</SCRIPT>


</td><td valign="center" nowrap align="middle" width="7%" class="a4"><%=rs1("toltopic")%>　</td><td valign="center" nowrap align="middle" width="7%" class="a3"><%=rs1("tolrestore")%>　</td><td width="16%" valign="center" class="a4"><table width="100%"><tr><td width="100%" height="38" align="right"><%=rs1("lasttime")%><br>by <a href=Profile.asp?username=<%=rs1("lastname")%>><%=rs1("lastname")%></a> <img src="images/lp.gif"></td></tr></table></td><td valign="center" align="center" width="7%" class="a3"><span style="CURSOR: hand" onclick="window.external.AddFavorite('<%=cluburl%>/Default.asp?id=<%=rs1("id")%>', '<%=rs1("bbsname")%>')"><img src="images/addfav.gif" border="0"></span> </td></tr>

<%
end sub


%>