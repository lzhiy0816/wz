<!-- #include file="setup.asp" -->
<%
forumid=int(Request("forumid"))

sql="select * from bbsconfig where id="&forumid&""
rs.Open sql,Conn

if rs.eof then
message=message&"<li>���ݿ��в�����IDΪ "&forumid&" ������"
error(""&message&"")
end if

bbsname=rs("bbsname")
moderated=rs("moderated")
logo=rs("logo")
pass=rs("pass")
rs.close
top
''''''''''��̳�涨��֤'''''''
if pass="0" then
error("<li>����̳��ʱ�رգ����ٽ��ܷ��ʣ�")
elseif pass="2" then
if Request.Cookies("username")=empty then error("<li>ֻ��<a href=login.asp>��¼</a>����������̳")
elseif pass="3" then
if Request.Cookies("username")=empty then error("<li>ֻ��<a href=login.asp>��¼</a>����������̳")
if membercode<2 and instr(moderated,"|"&Request.Cookies("username")&"|")<=0 then error("<li>ֻ�������α����ϵȼ����������̳")
end if
'''''''''''''''''''''''''''''
acturl="ShowForum.asp?forumid="&forumid&""
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<meta http-equiv=refresh content=300>
<script>
if(top==self)document.location='Default.asp?id=<%=forumid%>';
var key_word="<%=Request.Cookies("key_word")%>"
var cookieusername="<%=Request.Cookies("username")%>"
</script>

<title><%=bbsname%></title><table width="97%" align="center" border="0"><tr><td valign="top" width="30%">

<SCRIPT>if ("<%=logo%>"!=''){document.write("<img border=0 src=<%=logo%> onload='javascript:if(this.height>60)this.height=60;'")}else{document.write("<img border=0 src=images/lzylogo.gif>")}</SCRIPT>
</td><td valign="center" align="top">��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">��<a href="ShowForum.asp?forumid=<%=forumid%>"><%=bbsname%></a></td>
	<td align="right"><img src="images/jt.gif"> <a href="apply.asp">������̳</a><br>
	<img src="images/jt.gif"> <a href="supervise.asp?forumid=<%=forumid%>">��̳����</a></td></tr></table><br>
<!-- #include file="inc/line.asp" --><%
forumidonline=conn.execute("Select count(sessionid)from online where forumid="&forumid&"")(0)
regforumidonline=conn.execute("Select count(sessionid)from online where forumid="&forumid&" and username<>''")(0)




if request("order")<>"" then
order=server.htmlencode(Request("order"))
else
order="lasttime"
end if


if request("TimeLimit")<>"" then TimeLimit="and lasttime>"&SqlNowString&"-"&int(request("TimeLimit"))&""



if Request("search")="goodtopic" then search="and goodtopic=1 "



if Request.Cookies("pagesetup")=empty then
pagesetup=20   '�趨ÿҳ����ʾ����
else
pagesetup=int(Request.Cookies("pagesetup"))
if pagesetup > 30 then pagesetup=20
end if


topsql="where deltopic<>1 and forumid="&forumid&" "&search&" "&TimeLimit&" or toptopic=2"

count=conn.execute("Select count(id) from [forum] "&topsql&"")(0)
TotalPage=cint(count/pagesetup)  '��ҳ��
if TotalPage < count/pagesetup then TotalPage=TotalPage+1
PageCount = cint(Request.QueryString("ToPage"))
if PageCount < 1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage

sql="select * from [forum] "&topsql&" order by toptopic Desc,"&order&" Desc"
rs.Open sql,Conn,1
rs.pagesize=pagesetup   'ÿҳ��ʾ����
if TotalPage>0 then rs.absolutepage=PageCount '��ת��ָ��ҳ��
%>

<SCRIPT>valigntop()</SCRIPT>

<table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr><td width="93%" height="25" class="a1">��<img loaded="no" src="images/plus.gif" id="followImg0" style="cursor:hand;" onclick="loadThreadFollow(0,<%=forumid%>)"> Ŀǰ��̳������ <b><%=onlinemany%></b> �ˣ�������̳���� <b><%=forumidonline%></b> �����ߡ�����ע���û� <b><%=regforumidonline%></b> �ˣ��ÿ� <b><%=forumidonline-regforumidonline%></b> �ˡ�</td><td align="middle" width="7%" height="25" class="a1"><a href="javascript:this.location.reload()"><img src="images/refresh.gif" border="0"></a></td></tr><tr height="25" style="display:none" id="follow0"><td id="followTd0" align="left" class="a4" width="94%" colspan="5"><div onclick="loadThreadFollow(0,<%=forumid%>)"><table width="100%" cellpadding="10"><tr><td width="100%">Loading...</td></tr></table></div></td></tr></tr></table></td></tr></table>
<SCRIPT>valignbottom()</SCRIPT>
<br>
<table height="30" cellspacing="0" cellpadding="0" width="97%" align="center" border="0"><tr><td align="left" width="20%"><a href="newtopic.asp?forumid=<%=forumid%>"><img src="images/skins/<%=Request.Cookies("skins")%>/post.gif" border="0" alt="��������"></a></td><td align="right" width="80%"><img src="images/jt.gif"> ������:<%=Count%>��<img src="images/jt.gif"> <a href="ShowForum.asp?forumid=<%=forumid%>&search=goodtopic">��̳����</a>��<img src="images/team.gif"> 
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option>��̳����</option><option>-------</option>
<SCRIPT>
var moderated="<%=moderated%>"
var list= moderated.split ('|'); 
for(i=0;i<list.length;i=i+1) {
if (list[i] !=""){document.write("<option value=profile.asp?username="+list[i]+">"+list[i]+"</option>")}
}
</SCRIPT>
</select>
</td></tr></table><table height="28" cellspacing="1" cellpadding="1" width="97%" align="center" border="0"><tr><td align="middle" width="3%"><img height="18" src="images/announce.gif" width="18" align="middle" alt="��������"></td><td width="75%"><marquee onmouseover="this.stop()" onmouseout="this.start()" width="80%" scrollamount="3"><a style=cursor:hand onclick="javascript:open('affiche.asp','','width=400,height=180,resizable,scrollbars')"><%=affichetitle%></a>����<%=affichetime%>��</marquee></td><td width="22%" align="right"><select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option value="?forumid=<%=forumid%>">
�鿴���е�����</option><option value="?forumid=<%=forumid%>&TimeLimit=1">
�鿴һ���ڵ�����</option><option value="?forumid=<%=forumid%>&TimeLimit=2">
�鿴�����ڵ�����</option><option value="?forumid=<%=forumid%>&TimeLimit=7">
�鿴һ�����ڵ�����</option><option value="?forumid=<%=forumid%>&TimeLimit=15">
�鿴������ڵ�����</option><option value="?forumid=<%=forumid%>&TimeLimit=30">
�鿴һ�����ڵ�����</option><option value="?forumid=<%=forumid%>&TimeLimit=60">
�鿴�������ڵ�����</option><option value="?forumid=<%=forumid%>&TimeLimit=180">
�鿴�����ڵ�����</option></select></td></tr></table>


<SCRIPT>valigntop()</SCRIPT>
<table cellspacing="1" cellpadding="0" width="97%" align="center" border="0" class="a2"><tr height="25" id=TableTitleLink><td width="3%" class="a1">��</td><td width="3%" class="a1">��</td><td align="middle" height="24" class="a1" width="45%"><a href="ShowForum.asp?forumid=<%=forumid%>&order=topic&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">����</a></td><td align="middle" width="9%" height="24" class="a1"><a href="ShowForum.asp?forumid=<%=forumid%>&order=username&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">����</a></font></td><td align="middle" width="6%" height="24" class="a1"><a href="ShowForum.asp?forumid=<%=forumid%>&order=replies&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">�ظ�</a></td><td align="middle" width="7%" height="24" class="a1"><a href="ShowForum.asp?forumid=<%=forumid%>&order=Views&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">���</a></td><td width="27%" height="24" class="a1" align="center"><a href="ShowForum.asp?forumid=<%=forumid%>&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>">������</a></td></tr>
<%
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
if Not Response.IsClientConnected then responseend

newtopic=""
if rs("posttime")+1>now() then newtopic="<img src=images/new.gif>"

%><script>ShowForum("<%=rs("ID")%>","<%=rs("topic")%>","<%=newtopic%>","<%=rs("username")%>","<%=rs("Views")%>","<%=rs("icon")%>","<%=rs("toptopic")%>","<%=rs("locktopic")%>","<%=rs("pollresult")%>","<%=rs("goodtopic")%>","<%=rs("replies")%>","<%=rs("lastname")%>","<%=rs("lasttime")%>");</script>
<%
RS.MoveNext
loop
RS.Close
%></table>
<SCRIPT>valignbottom()</SCRIPT>
<table cellspacing="1" cellpadding="1" width="97%" align="center" border="0"><tr height="25"><td width="100%" height="2"><table cellspacing="0" cellpadding="3" width="100%" border="0"><tr><td height="2">
		<b>����̳���� <font color=990000><%=TotalPage%></font> ҳ [ 


<b>
<script>
PageCount=<%=TotalPage%>
topage=<%=PageCount%>
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?topage="+i+"&forumid=<%=forumid%>&order=<%=Request("order")%>&search=<%=Request("search")%>&TimeLimit=<%=Request("TimeLimit")%>>"+ i +"</a> ");
}
}
}
</script>

</b>]

</b></td><form name="form" action="searchok.asp?forumid=<%=forumid%>&search=key&searchxm2=topic" method="post"><td height="2" align="right">����������<input name="content" value="����ؼ���" onfocus="this.value = &quot;&quot;;" style="font-family: Tahoma,Verdana,����; font-size: 12px; line-height: 15px" size="20"> <input type="submit" value="����" name="submit" style="font-family: Tahoma,Verdana,����; font-size: 12px; line-height: 15px"> </td></form></tr></table></td></tr></table><br><center><table cellspacing="4" cellpadding="0" width="80%" border="0"><tr><td nowrap width="200"><img alt="" src="images/f_new.gif" border="0"> ������ (�лظ�������)</td><td nowrap width="100"><img alt="" src="images/f_hot.gif" border="0"> �������� </td><td nowrap width="100"><img alt="" src="images/f_locked.gif" border="0"> �ر�����</td><td nowrap width="150"><img src="images/topicgood.gif"> ��������</td></tr><tr><td nowrap width="200"><img alt="" src="images/f_norm.gif" border="0"> ������ (û�лظ�������)</td><td nowrap width="100"><img alt="" src="images/f_poll.gif" border="0"> ͶƱ����</td><td nowrap width="100"><img alt="" src="images/f_top.gif" border="0"> �ö�����</td><td nowrap width="150"><img src="images/my.gif"> �Լ����������</td></tr></table></center></div></td></tr></table>
<iframe height=0 width=0 name=hiddenframe></iframe>
<%
htmlend
%>