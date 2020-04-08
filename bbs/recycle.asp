<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


if membercode < 4 then
error2("您的权限不够，无法进入！")
end if

select case Request("menu")

case "deoxidize"
for each ho in request.form("id")
conn.execute("update [forum] set deltopic=0,lasttime='"&now()&"',lastname='"&Request.Cookies("username")&"' where id="&ho&" and deltopic=1")
next


case "clean"

if ""&Request("day")&"" < "2" and membercode < 5 then error2("只有社区区长才有权限清空 1 天前的帖子")

conn.execute("delete from [forum] where deltopic=1 and lasttime<"&SqlNowString&"-"&Request("day")&"")
log("清空回收站内 "&Request("day")&" 天以前的文章！")
error2("已经清空回收站内 "&Request("day")&" 天以前的文章！")
end select



sql="select * from forum where deltopic=1 order by lasttime Desc"
rs.Open sql,Conn,1

pagesetup=20 '设定每页的显示数量
rs.pagesize=pagesetup
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>







<title>回收站</title>

<script>
function CheckAll(form){for (var i=0;i<form.elements.length;i++){var e = form.elements[i];if (e.name != 'chkall')e.checked = form.chkall.checked;}}
</script>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<a href="recycle.asp">回收站</a></td>
</tr>
</table>


<br>
<SCRIPT>valigntop()</SCRIPT>
<TABLE cellSpacing=1 cellPadding=1 width=97% align=center border=0 class=a2>
  <form method="POST" action=recycle.asp?menu=deoxidize>
<TR height=25>
      <TD align=middle width=3% class=a1>　</TD>
      <TD align=middle width=3% class=a1>
<input type=checkbox name=chkall onclick=CheckAll(this.form) value="ON"></TD>
      <TD width="46%" height=24 align=middle class=a1>主题</TD>
      <TD align=middle width=9% height=24 class=a1>作者</TD>
      <TD align=middle width=5% height=24 class=a1>回复</TD>
      <TD align=middle width=4% height=24 class=a1>点击</TD>
      <TD width=30% height=24 class=a1>　　 删除时间　　　　 删除者</TD>
    </TR>


<%

i=0
Do While Not RS.EOF and i<pagesetup
i=i+1




if rs("toptopic")=2 then
reimage="<img src=images/top.gif>"
elseif rs("toptopic")=1 then
reimage="<img src=images/f_top.gif>"
elseif rs("locktopic")=1 then
reimage="<img src=images/f_locked.gif>"
elseif rs("polltopic")<>"" then
reimage="<img src=images/f_poll.gif>"
elseif rs("replies")>15 or rs("views") > 150 then
reimage="<img src=images/f_hot.gif>"
elseif rs("replies")>0 then
reimage="<img src=images/f_new.gif>"
else
reimage="<img src=images/f_norm.gif>"
end if


if rs("replies")=0 then
replies="-"
else
replies=rs("replies")
end if

%>

<TR height=25>
      <TD align=middle width=3% class=a4 height=24><%=reimage%></TD>
      <TD align=middle width=3% class=a3>
<INPUT type=checkbox value=<%=rs("id")%> name=id></TD>
      <TD width="46%" align=left class=a4>&nbsp;<img border=0 src=images/brow/<%=rs("icon")%>.gif> <a onclick=min_yuzi() target=message href=ShowPost.asp?id=<%=rs("id")%>><%=rs("topic")%></a></TD>
      <TD align=middle width=9% class=a3><a href=Profile.asp?username=<%=rs("username")%>><%=rs("username")%>　</TD>
      <TD align=middle width=5% class=a4><%=replies%>　</TD>
      <TD align=middle width=4% class=a3><%=rs("views")%>　</TD>
      <TD align=left width=30% class=a4>&nbsp;<%=rs("lasttime")%> | <a href=Profile.asp?username=<%=rs("lastname")%>><%=rs("lastname")%></a></TD>
    </TR>
    

<%
RS.MoveNext
loop
RS.Close
%>







</TABLE>
<SCRIPT>valignbottom()</SCRIPT>
<TABLE cellSpacing=0 cellPadding=1 width=97% align=center border=0>
  <TBODY>
<TR height=25>
      <TD width="100%" height=2> 
        <TABLE cellSpacing=0 cellPadding=3 width="100%" border=0>
<TBODY>
<TR>
<TD height=2>
<input type=submit onclick="{if(confirm('您确定要还原所选的主题?')){return true;}return false;}" value="还 原">
<b>共有
<font color="990000"><%=TotalPage%></font> 页 [
<script>
PageCount=<%=TotalPage%> //总页数
topage=<%=PageCount%>   //当前停留页
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?topage="+i+">"+ i +"</a> ");
}
}
}
</script>]</b></TD>
<TD align=right height=2>
</form>
 <form method="POST" action=?menu=clean>

清空 <input maxLength="1" size="1" value="7" name=day> 天以前的文章 <INPUT onclick="{if(confirm('您确定要清空回收站?')){return true;}return false;}" type=submit value="确定"> 

 </form>




</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>




<br>



<center>






<p></p>







<div align="center">
<center>







<table cellSpacing="4" cellPadding="0" width="80%" border="0">
<tr>
<td noWrap width="200">
<img alt src="images/f_new.gif" border="0">&nbsp;打开主题 (有回复的主题)</td>
<td noWrap width="100">
<img alt src="images/f_hot.gif" border="0">&nbsp;热门主题 </td>
<td noWrap width="100">
<img alt src="images/f_locked.gif" border="0">&nbsp;关闭主题</td>
<td noWrap width="150">
<img src="images/topicgood.gif">
精华主题</td>
</tr>
<tr>
<td noWrap width="200">
<img alt src="images/f_norm.gif" border="0">&nbsp;打开主题 (没有回复的主题)</td>
<td noWrap width="100">
<img alt src="images/f_poll.gif" border="0">&nbsp;投票主题</td>
<td noWrap width="100">
<img alt src="images/f_top.gif" border="0">&nbsp;置顶主题</td>
<td noWrap width="150">
<img src="images/my.gif">
自己发表的主题</td>
</tr>
</table>


</center>
</div>




<%


htmlend
%>