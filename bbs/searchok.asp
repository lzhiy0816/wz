<!-- #include file="setup.asp" -->
<%

if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

if instr(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&"") = 0 then error("<li>来源错误")

top

search=Request("search")
forumid=Request("forumid")
content=server.htmlencode(Request("content"))
searchxm=Request("searchxm")
searchxm2=Request("searchxm2")
searchxm2=replace(searchxm2,"@","&")

if content=empty then
content=Request.Cookies("username")
end if


if isnumeric(""&forumid&"") then
forumidor="forumid="&forumid&" and"
end if

if Request("search")="author" then
item=""&searchxm&"='"&content&"'"
elseif Request("search")="key" then
item=""&searchxm2&" like '%"&content&"%'"
end if

if request("TimeLimit")<>"" then TimeLimit="and lasttime>"&SqlNowString&"-"&int(request("TimeLimit"))&""


sql="select * from forum where deltopic<>1 and "&forumidor&" "&item&" "&TimeLimit&" order by lasttime Desc  "
rs.Open sql,Conn,1

count=rs.recordcount    '数据总条数
if Count=0 then
error("<li>对不起，没有找到您要查询的内容")
end if

pagesetup=20   '设定每页的显示数量
rs.pagesize=pagesetup   '每页显示条数
TotalPage=rs.pagecount  '总页数
PageCount = cint(Request.QueryString("ToPage"))
if PageCount <1 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
if TotalPage>0 then rs.absolutepage=PageCount '跳转到指定页数


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<script>
if(top==self)document.location='Default.asp';
var key_word="<%=content%>"
var cookieusername="<%=Request.Cookies("username")%>"
</script>




<title>搜索结果</title>



<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　搜索结果</td>
</tr>
</table>


　<table width=97% align=center>
<tr>
<td width="100%" align="right">搜索关键字：<font color="FF0000"><%=content%></font>　　共找到了 <b><font color="FF0000"><%=Count%></font></b> 篇相关帖子</td>
</tr>
</table>

<SCRIPT>valigntop()</SCRIPT>
    <table  cellSpacing="1" cellPadding="0" width=97% align="center" border="0" class=a2>
  <tr height="25">
    <td width="3%" class=a1>　</td>
    <td width="3%" class=a1></td>
    <td align="middle" height="24" class=a1 width="45%"> 主题</td>
    <td align="middle" width="9%" height="24" class=a1> 作者</td>
    <td align="middle" width="6%" height="24" class=a1> 回复</td>
    <td align="middle" width="7%" height="24" class=a1> 点击</td>
    <td width="27%" height="24" class=a1 align="center">最后更新</td>
  </tr>

<%

i=0
Do While Not RS.EOF and i<pagesetup
i=i+1


newtopic=""
if rs("posttime")+1>now() then newtopic="<img src=images/new.gif>"

%><script>ShowForum("<%=rs("ID")%>","<%=rs("topic")%>","<%=newtopic%>","<%=rs("username")%>","<%=rs("Views")%>","<%=rs("icon")%>","<%=rs("toptopic")%>","<%=rs("locktopic")%>","<%=rs("pollresult")%>","<%=rs("goodtopic")%>","<%=rs("replies")%>","<%=rs("lastname")%>","<%=rs("lasttime")%>");</script>
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
<TD height=2><b>&nbsp;本论坛共有
<font color="990000"><%=TotalPage%></font> 页 [

<%
TimeLimit=Request("TimeLimit")
searchxm2=replace(searchxm2,"&","@")

%>


<script>
PageCount=<%=TotalPage%> //总页数
topage=<%=PageCount%>   //当前停留页
for (var i=1; i <= PageCount; i++) {
if (i <= topage+3 && i >= topage-3 || i==1 || i==PageCount){
if (i > topage+4 || i < topage-2 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write(" "+ i +" ");}
else{
document.write("<a href=?topage="+i+"&forumid=<%=forumid%>&search=<%=search%>&searchxm=<%=searchxm%>&searchxm2=<%=searchxm2%>&content=<%=content%>&TimeLimit=<%=TimeLimit%>>"+ i +"</a> ");
}
}
}
</script>


]</b></TD>

              <form name="form" action="searchok.asp?forumid=<%=forumid%>&search=key&searchxm2=topic" method="post">
                <td height="2" align="right">
                快速搜索：<input name="content" style="font-family: Tahoma,Verdana,宋体; font-size: 12px; line-height: 15px" size="20">&nbsp;<input type="submit" value="搜索" name="submit" style="font-family: Tahoma,Verdana,宋体; font-size: 12px; line-height: 15px"> </td>
              </form>





</TR></TBODY></TABLE></TD></TR></TBODY></TABLE>




<br>



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


<IFRAME HEIGHT=0 WIDTH=0 NAME=hiddenframe></IFRAME>

<%


htmlend
%>