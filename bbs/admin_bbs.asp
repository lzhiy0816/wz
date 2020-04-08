<!-- #include file="setup.asp" -->
<%
if adminpassword<>session("pass") then response.redirect "admin.asp?menu=login"
id=server.htmlencode(Request("id"))


log(""&Request.ServerVariables("script_name")&"<br>"&Request.ServerVariables("Query_String")&"<br>"&Request.form&"")


%>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<br><center>
<p></p>

<%


select case Request("menu")
case "applymanage"
applymanage


case "bbsmanage"
bbsmanage

case "bbsmanagexiu"
bbsmanagexiu


case "bbsmanagexiuok"
bbsmanagexiuok

case "bbsadd"
bbsadd

case "bbsaddok"
bbsaddok

case "classs"
classs

case "classxiu"
classxiu

case "upclubconfig"
upclubconfig

case "upclubconfigok"
upclubconfigok

case "classadd"
if id="" then id=1
if not isnumeric(id) then error2("序号请用数字填写")
if Request("classname")="" then error2("请填写类别名称")
conn.Execute("insert into class(id,classname) values ('"&id&"','"&Request("classname")&"')")
classs

case "classdel"
conn.execute("delete from [class] where id="&id&"")
classs

case "classxiuok"
conn.execute("update [class] set classname='"&Request("classname")&"',id="&id&" where id="&Request("oldid")&"")
classs

case "bbsmanagedel"
conn.execute("delete from [bbsconfig] where id="&id&"")
error2("已经将该论坛的所有数据删除了！")

case "delforumok"

if request("jh")="1" then
jh=" and goodtopic<>1"
end if

if request("bbsid")<>"" then
bbsid="and forumid="&request("bbsid")&""
end if


conn.execute("delete from [forum] where lasttime<"&SqlNowString&"-"&request("TimeLimit")&" "&jh&" "&bbsid&"")

error2("已经将"&request("TimeLimit")&"天没有用户回复的主题删除了！")


case "delretopicok"
conn.execute("delete from [reforum] where posttime<"&SqlNowString&"-"&request("TimeLimit")&"")
error2("已经将"&request("TimeLimit")&"天的回帖删除了！")


case "delbbsconfigok"

if request("hide")="1" then
hide=" and hide=1"
end if

conn.execute("delete from [bbsconfig] where lasttime<"&SqlNowString&"-"&request("TimeLimit")&" "&hide&"")
error2("已经将"&request("TimeLimit")&"天没有新帖子的论坛删除了！")



case "deltopicok"
if request("topic")="" then
error2("您没有输入字符！")
end if
conn.execute("delete from [forum] where topic like '%"&request("topic")&"%' ")
error2("已经将标题里包含有 "&request("topic")&" 的帖子全部删除了！")


case "uniteok"
if Request("hbbs") = Request("ybbs") then
error2("不能选择相同论坛！")
end if
conn.execute("update [forum] set forumid="&Request("hbbs")&" where forumid="&Request("ybbs")&"")
error2("已经成功将2个论坛的资料合并了！")



end select

sub applymanage
%>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>

<table cellspacing="1" cellpadding="2" width="97%" border="0" class="a2" align="center">



  <tr class=a1 id=TableTitleLink>
<td align="center" height="25"><a href="?menu=applymanage&fashion=id">
ID</a></td>
<td width="20%" align="center" height="25">
<a href="?menu=applymanage&fashion=bbsname">论坛</a></td>
<td align=center height="25">版主</td>
<td align=center height="25"><a href="?menu=applymanage&fashion=toltopic">
主题</a></td>
<td align=center height="25"><a href="?menu=applymanage&fashion=tolrestore">
帖子</a></td>

<td align="center" height="25">操作</td>
    
  </tr>


<%

if Request("fashion")=empty then
fashion="tolrestore"
else
fashion=Request("fashion")
end if


sql="select * from bbsconfig where hide=1 order by "&fashion&" Desc"
rs.Open sql,Conn,1

pagesetup=20 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1

%>
  <tr class=a3>
<td align="center" height="25"><%=rs("id")%></td>
<td><a target=_blank href=ShowForum.htm?forumid=<%=rs("id")%>><%=rs("bbsname")%></a></td>
<td align=center><%=rs("moderated")%></td>
<td align=center><b><font color=red><%=rs("toltopic")%></font></b></td>
<td align=center><b><font color=red><%=rs("tolrestore")%></font></b></td>

<td align="center"><a href=?menu=bbsmanagexiu&id=<%=rs("id")%>>编辑论坛</a> | <a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsmanagedel&id=<%=rs("id")%>>删除论坛</a>
　</td>
    
  </tr>
<%
RS.MoveNext
loop
RS.Close

%>
</table>

[<b>
<script>
PageCount=<%=TotalPage%> //总页数
topage=<%=PageCount%>   //当前停留页
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?menu=applymanage&fashion=<%=Request("fashion")%>&topage="+i+">"+ i +"</a> ");
}
}
}
</script>

</b>]
<%

end sub


sub classs
%>
<script>function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}</script>
<table border="0" width="90%">
	<tr>
		<td align="center"><form method="post" action="?menu=classadd">
类别名称：（例如：电脑网络）<INPUT size=20 name=classname> 序号：<INPUT size=2 name=id value=<%=conn.execute("Select max(id)from class")(0)+1%>>
<input type=submit value="建立"></form></td>
		<td align="center"><form method="post" action="?menu=bbsmanagexiu">
查找论坛：<INPUT size=2 name=id value="ID" onfocus="this.value = ''" >
<input type=submit value="确定"></form></td>
	</tr>
</table>





<%


sql="select * from class order by id"
rs.Open sql,Conn
do while not rs.eof

%>


<table border="0" cellpadding="5" cellspacing="0"  width="90%" class=a2>
  <tr class=a1 id=TableTitleLink>
    <td width="50%">
<%=rs("id")%>.<%=rs("classname")%>
    </td>

    <td width="50%" align="right">
    <a href=?menu=bbsadd&id=<%=rs("id")%>>建立论坛</a> | <a href=?menu=classxiu&id=<%=rs("id")%>>
    编辑分类</a> | <a href=?menu=classdel&id=<%=rs("id")%>>
    删除分类</a>

    　</td>
  </tr>
</table>
<%


sql="select * from bbsconfig  where classid="&rs("id")&" and hide=0"
rs1.Open sql,Conn
do while not rs1.eof
%>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="90%" class=a4>
  <tr>

<td width="50%">

<li><%=rs1("bbsname")%></td>
    
<td width="50%" align="right"><a href=?menu=bbsmanagexiu&id=<%=rs1("id")%>>编辑论坛</a> | <a onclick=checkclick('您确定要删除该论坛的所有资料?') href=?menu=bbsmanagedel&id=<%=rs1("id")%>>删除论坛</a>
　</td>
    
  </tr>
</table>
<%
rs1.movenext
loop
rs1.close




rs.movenext
loop
rs.close

end sub

sub classxiu
sql="select * from class where id="&id&""
rs.Open sql,Conn
%>
<form method="post" action="?menu=classxiuok">
<input type=hidden name=oldid  value=<%=id%>>
类别名称：<INPUT size=20 name=classname value=<%=rs("classname")%>> 序号：<INPUT size=2 name=id value=<%=id%>>
<input type=submit value="修改"><center><br><br><a href=javascript:history.back()>< 返 回 ></a>
<%
rs.close
end sub

sub bbsadd

%>


<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">建立论坛资料</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

<form name="form" method="post" action="?menu=bbsaddok">
<input type=hidden name=classid value=<%=id%>>

论坛名称</td>
    <td class=a3>

<input size="20" name="bbsname"></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛版主</td>
    <td class=a3>

<input size="30" name=moderated> 多版主添加请用|分隔，如：yuzi|裕裕
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛介绍</td>
    <td class=a3>

<textarea rows="5" name="intro" cols="50"></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛状态</td>
    <td class=a3>

<select size="1" name="pass">
<option value=0>关闭论坛</option>
<option value=1 selected>正规论坛</option>
<option value=2>会员论坛</option>
<option value=3>嘉宾论坛</option>
</select>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

小图标URL</td>
    <td class=a3>

<input size="30" name="icon"> 显示在社区首页论坛介绍左边
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

大图标URL</td>
    <td class=a3>

<input size="30" name="logo"> 显示在论坛左上角
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>

是否显示在论坛列表　　　　　　　　　　　　</td>
    <td class=a3>

<input type="radio" CHECKED value="0" name="hide" id=hide1><label for=hide1>显示</label>
 
<input type="radio" value="1" name="hide" id=hide2><label for=hide2>隐藏</label> 　</td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">

　 <input type="submit" value=" 建 立 " name="Submit"><br></td></tr></table>
</form>
<center><br><a href=javascript:history.back()>< 返 回 ></a>


<%
end sub
sub bbsaddok
if Request("bbsname")="" then error2("请输入论坛名称")


temp="|"&Request("moderated")&"|"
temp=replace(temp,"||","|")
master=split(temp,"|")
for i = 1 to ubound(master)-1
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"的用户资料还未注册")
next



rs.Open "bbsconfig",Conn,1,3
rs.addnew
rs("classid")=Request("classid")
rs("bbsname")=server.htmlencode(Request("bbsname"))
rs("moderated")=temp
rs("pass")=Request("pass")
rs("intro")=server.htmlencode(Request("intro"))
rs("icon")=server.htmlencode(Request("icon"))
rs("logo")=server.htmlencode(Request("logo"))
rs.update
id=rs("id")

rs.close

classs

end sub




sub bbsmanagexiuok

if Request("bbsname")="" then error2("请输入论坛名称")

if Request("hide")=0 and Request("classid")="" then
error2("请指定论坛的类别")
end if


temp="|"&Request("moderated")&"|"
temp=replace(temp,"||","|")
master=split(temp,"|")
for i = 1 to ubound(master)-1
If conn.Execute("Select id From [user] where username='"&HTMLEncode(master(i))&"'" ).eof and master(i)<>"" Then error2(""&master(i)&"的用户资料还未注册")
next


sql="select * from bbsconfig where id="&id&""
rs.Open sql,Conn,1,3


if Request("classid")<>empty then rs("classid")=Request("classid")
rs("bbsname")=server.htmlencode(Request("bbsname"))
rs("moderated")=temp
rs("pass")=Request("pass")
rs("intro")=server.htmlencode(Request("intro"))
rs("icon")=server.htmlencode(Request("icon"))
rs("logo")=server.htmlencode(Request("logo"))
rs("hide")=Request("hide")
rs.update

rs.close
%>
编辑成功<br><br><a href=javascript:history.back()>返 回</a>
<%
end sub


sub bbsmanagexiu
rs.open "class order by id",conn
do while not rs.eof
classlist=""+classlist+"<option value="&rs("id")&">"&rs("classname")&"</option>"
rs.movenext
loop
rs.close
sql="select * from bbsconfig where id="&id&""
rs.Open sql,Conn
%>


<form method="post" action="?menu=bbsmanagexiuok"><input type=hidden name=id value=<%=rs("id")%>>
<table cellspacing="1" cellpadding="2" width="80%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan="2">编辑论坛资料</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>

论坛名称</td>
    <td class=a3>


<input size="10" name="bbsname" value="<%=rs("bbsname")%>"> 类别 
<select name="classid">
<option value="<%=rs("classid")%>">默认</option>
<%=classlist%>
</select></td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛版主</td>
    <td class=a3>


<input size="30" name=moderated value="<%=rs("moderated")%>"> 多版主添加请用|分隔，如：yuzi|裕裕
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛介绍</td>
    <td class=a3>


<textarea rows="5" name="intro" cols="50"><%=rs("intro")%></textarea></td></tr>
   <tr height=25>
    <td class=a3 align=middle>


论坛状态</td>
    <td class=a3>


<select size="1" name="pass">
<option value=0 <%if rs("pass")=0 then%>selected<%end if%>>关闭论坛</option>
<option value=1 <%if rs("pass")=1 then%>selected<%end if%>>正规论坛</option>
<option value=2 <%if rs("pass")=2 then%>selected<%end if%>>会员论坛</option>
<option value=3 <%if rs("pass")=3 then%>selected<%end if%>>嘉宾论坛</option>
</select>
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


小图标URL</td>
    <td class=a3>


<input size="30" name="icon" value="<%=rs("icon")%>"> 显示在社区首页论坛介绍左边
</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


大图标URL</td>
    <td class=a3>


<input size="30" name="logo" value="<%=rs("logo")%>"> 显示在论坛左上角</td></tr>
   <tr height=25>
    <td class=a3 align=middle>


是否显示在论坛列表</td>
    <td class=a3>


<input type="radio" <%if rs("hide")=0 then%>CHECKED <%end if%>value="0" name="hide" value="0" id=hide1><label for=hide1>显示</label> 
<input type="radio" <%if rs("hide")=1 then%>CHECKED <%end if%>value="1" name="hide" value="1" id=hide2><label for=hide2>隐藏</label> </td></tr>
   <tr height=25>
    <td class=a3 align=middle colspan="2">


<input type="submit" value=" 编 辑 " name="Submit1"></td></tr></table><br></form>
<center><br><a href=javascript:history.back()>< 返 回 ></a>
<%
end sub



sub bbsmanage

%>

论坛数：<b><font color=red><%=conn.execute("Select count(id)from bbsconfig")(0)%></font></b>　　主题数：<b><font color=red><%=conn.execute("Select count(id)from forum")(0)%></font></b>　　回帖数：<b><font color=red><%=conn.execute("Select count(id)from reforum")(0)%></font></b><br>

<form method="post" action="?menu=delforumok">

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle>批量删除主题</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle>
    
删除 <select name=TimeLimit>
<option value="90">90</option>
<option value="180" selected>180</option>
<option value="360">360</option>
</select> 天没有用户回复的主题　<input type="checkbox" value="1" name="jh" id=jh checked><label for="jh">精华帖子除外</label>
<br>
<select name="bbsid">
<option value="">所有论坛</option>

<%
sql="select id,bbsname,classid from bbsconfig where hide=0 order by classid,id"
rs.Open sql,Conn
do while not rs.eof
Classid=Trim(Rs("classid"))
		if TClass <> Classid Then
		Response.write "<option style=BACKGROUND-COLOR:ECF5FF value=''>╋ "&Conn.Execute("Select classname From class where id="&Classid)(0)&"</option>"
		TClass = Classid
		end if
		
	Response.write "<option value="&rs("id")&""&selected&">　├『"&rs("bbsname")&"』</option>"
	rs.movenext
loop
rs.close
%>

</select>
 <input type="submit" value=" 确 定 ">
 


</td></form></tr>
   <tr height=25>
    <td class=a4 align=middle>
    <form method="post" action="?menu=deltopicok">
删除标题里包含有 <input size="20" name="topic"> 的所有帖子 <input type="submit" value="确定">

　</td></tr></table></form>


<form method="post" action="?menu=delretopicok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>批量删除回帖</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>


删除 <select name=TimeLimit>
<option value="90">90</option>
<option value="180" selected>180</option>
<option value="360">360</option>
</select> 天以前的回帖  <input type="submit" value=" 确 定 ">
</td></tr></table></form>

<form method="post" action="?menu=delbbsconfigok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>批量删除论坛</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
删除 <select name=TimeLimit>
<option value="30">30</option>
<option value="60" selected>60</option>
<option value="90">90</option>
</select> 天没有新帖子的论坛 <input type="checkbox" value="1" name="hide" id=hide checked><label for="hide">列表中的论坛除外</label><br><input type="submit" value=" 确 定 ">
</td></tr></table></form>



<form method="post" action="?menu=uniteok">
<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>合并论坛数据</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>

将 <select name="ybbs">
<%

rs.Open "bbsconfig where hide=0",Conn

do while not rs.eof
%>
<option value="<%=rs("id")%>">『<%=rs("bbsname")%>』</option>
<%
rs.movenext
loop
rs.close
%>
</select> 合并到 <select name="hbbs">
<%
rs.Open "bbsconfig where hide=0",Conn
do while not rs.eof
%>
<option value="<%=rs("id")%>">『<%=rs("bbsname")%>』</option>
<%
rs.movenext
loop
rs.close
%>
</select>
<br>
<INPUT type=submit value=" 确 定 " name=submit>
</td></tr></table></form>






<%
end sub

sub upclubconfig
%>

<table cellspacing="1" cellpadding="2" width="70%" border="0" class="a2" align="center">
  <tr height=25>
    <td class=a1 align=middle colspan=2>社区资料更新</td>
  </tr>
   <tr height=25>
    <td class=a3 align=middle colspan=2>
    
此操作将更新论坛资料、用户资料<br>（防止论坛统计错误和用户类型显示不正确的情况发生）<br>
<a href="?menu=upclubconfigok">点击这里更新论坛统计数据</a><br>
</td></tr></table><br>

<%
end sub

sub upclubconfigok

rs.Open "bbsconfig",Conn
do while not rs.eof

allarticle=conn.execute("Select count(forumid) from forum where forumid="&rs("id")&"")(0)
if allarticle>0 then
allrearticle=conn.execute("Select sum(replies) from forum where forumid="&rs("id")&"")(0)
else
allrearticle=0
end if

conn.execute("update [bbsconfig] set toltopic="&allarticle&",tolrestore="&allarticle+allrearticle&" where ID="&rs("id")&"")


rs.movenext
loop
rs.close
%>
操作成功！<br>
<br>
已经更新社区所有论坛的统计数据<br>

<%
end sub

htmlend



%>