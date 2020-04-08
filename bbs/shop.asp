<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")


content=replace(server.htmlencode(Request("content")), "'", "&#39;")

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3


select case Request("menu")

case "experience"

if Request("how")<1 then
message=message&"<li>数量不能小于1"
error(message)
end if

if rs("money")<300*Request("how") then
message=message&"<li>您的金币不够！"
error(message)
end if
Randomize
d1=fix(rnd*300)+1
rs("experience")=rs("experience")+d1*Request("how")
rs("money")=rs("money")-300*Request("how")
rs.update
rs.close
error2("已经增加了 "&d1*Request("how")&" 点经验值！")

case "lookip"
if rs("money")<500 then
error("<li>您的金币不够！")
end if
rs("money")=rs("money")-500
rs.update
rs.close
Response.Cookies("lookip")=1
response.redirect "online.asp"


case "colony"
if content="" then
error2("请填写群发内容!")
end if
if rs("money")<1000 then
error("<li>您的金币不够！")
end if
rs("money")=rs("money")-1000
rs.update
rs.close
sql="select username from online where username<>''"
rs.Open sql,Conn
do while not rs.eof
Count=Count+1


conn.Execute("insert into message (author,incept,content) values ('"&Request.Cookies("username")&"','"&rs("username")&"','【会员广播】："&content&"')")

conn.execute("update [user] set newmessage=newmessage+1 where username='"&rs("username")&"'")
rs.movenext
loop
rs.close
error2("发送成功！\n\n共发送给 "&Count&" 位在线会员")



case "thew"
if rs("money")<100 then
message=message&"<li>您的金币不够！"
error(message)
end if
if rs("userlife")=100 then
message=message&"<li>体力已满，无需增加！"
error(message)
end if
rs("money")=rs("money")-100
rs("userlife")=100
rs.update
rs.close
error2("您的体力已经全满！")

case "flowers"
vs=server.htmlencode(Trim(Request("vs")))
if vs=Request.Cookies("username") then error("<li>不能自己送自己！")

If conn.Execute("Select id From [user] where username='"&vs&"'" ).eof Then
error("<li>系统不存在"&vs&"的资料")
end if

if rs("money")<50 then
error("<li>您的金币不够！")
end if

rs("money")=rs("money")-50
rs.update
rs.close

Randomize
d1=fix(rnd*50)+1

conn.execute("update [user] set experience=experience+"&d1&" where username='"&vs&"'")

sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&vs&"','【系统消息】："&Request.Cookies("username")&"送您1束鲜花，您增加了"&d1&"点经验值！')"
conn.Execute(SQL)
conn.execute("update [user] set newmessage=newmessage+1 where username='"&vs&"'")

error2(""&vs&"已经增加了"&d1&"点经验值！")



case "egg"
vs=server.htmlencode(Trim(Request("vs")))
if vs=Request.Cookies("username") then error("<li>不能自己送自己！")



If conn.Execute("Select id From [user] where username='"&vs&"'" ).eof Then
error("<li>系统不存在"&vs&"的资料")
end if
If conn.Execute("Select experience From [user] where username='"&vs&"'" )(0)<50 Then error("<li>对方经验值少于50，您不能再向他（她）丢鸡蛋了")


if rs("money")<50 then
error("<li>您的金币不够！")
end if

rs("money")=rs("money")-50
rs.update
rs.close

Randomize
d1=fix(rnd*50)+1

conn.execute("update [user] set experience=experience-"&d1&" where username='"&vs&"'")

sql="insert into message(author,incept,content) values ('"&Request.Cookies("username")&"','"&vs&"','【系统消息】："&Request.Cookies("username")&"送您1粒鸡蛋，您减少了"&d1&"点经验值！')"
conn.Execute(SQL)
conn.execute("update [user] set newmessage=newmessage+1 where username='"&vs&"'")


error2(""&vs&"已经减少了"&d1&"点经验值！")

end select

%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">　<a href="shop.asp">社区商店</a></font></td>
</tr>
</table>
<br>
<title>社区商场</title>

<SCRIPT>
function colony(){
var id=prompt("请输入你要群发的讯息！","");if(id){document.location='?menu=colony&content='+id+'';}
}
</SCRIPT>
<center>



<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="111111" width="80%">
<tr>
<td width="100%">
<div align="center">
<center>
<table border="0" cellpadding="3" cellspacing="1" width="90%" class=a2>
<tr class=a1>
<td width="50%" height="5" align="center" colspan="2"><b>
体力药丸</b></td>
<td width="50%" height="5" align="center" colspan="2"><b>
经验药丸</b></td>
</tr>
<tr>
<td width="15%" height="4" align="center" bgcolor="FFFFFF" rowspan="2"><FONT color=cccccc>
<IMG src="images/plus/9.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：体力药丸<br>
价格：100 <b>金币</b><br>
功能：快速补充体力</td>
<td width="15%" height="4" align="center" bgcolor="FFFFFF" rowspan="2"><FONT color=cccccc>
<IMG src="images/plus/7.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">
  名称：经验药丸<br>
  价格：300 <b>金币</b><br>
  功能：提升经验</td>
</tr>
<tr>
<td width="25%" height="2" bgcolor="FFFFFF"><form method=POST action=?menu=thew>
<input type="submit" value="我要购买" name="Submit"></td></form>
<td width="25%" height="2" bgcolor="FFFFFF"><form method=POST action=?menu=experience>
数量：<input maxLength="2" size="2" name="how" value="1">
<input type="submit" value="我要购买" name="Submit"></td></form>
</tr>
</table>
</center>
</div><br>
</td>
</tr>


<tr>
<td width="100%" height="97"><div align="center">
<center>
<table border="0" cellpadding="3" cellspacing="1" width="90%" class=a2>
<tr class=a1>
<td width="50%" height="5" align="center" colspan="2"><b>
鲜</b>　<b>花</b></td>
<td width="50%" height="5" align="center" colspan="2"><b>
鸡</b>　<b>蛋</b></td>
</tr>
<tr>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/flowers.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：鲜花<br>
价格：50 <b>金币</b><br>
功能：增加对方的经验值</td>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/egg.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：鸡蛋<br>
价格：50 <b>金币</b><br>
功能：降低对方的经验值</td>
</tr>
<tr>
<td width="25%" height="2" bgcolor="FFFFFF"><form method=POST action=?menu=flowers>
<input size="9" name="vs" value="对方用户名" onfocus="this.value = &quot;&quot;;">
<input type="submit" value=" 确 定 " name="Submit1"></td></form>
<td width="25%" height="2" bgcolor="FFFFFF"><form method=POST action=?menu=egg>
<input size="9" name="vs" value="对方用户名" onfocus="this.value = &quot;&quot;;">
<input type="submit" value=" 确 定 " name="Submit2"></td></form>
</tr>
</table>
</center>
</div><br>
</td>
</tr>


<tr>
<td width="100%" height="97"><div align="center">
<center>
<table border="0" cellpadding="3" cellspacing="1" width="90%" class=a2>
<tr class=a1>
<td width="50%" height="5" align="center" colspan="2"><b>
ＩＰ查看器</b></td>
<td width="50%" height="5" align="center" colspan="2"><b>
讯息群发器</b></td>
</tr>
<tr>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/lookip.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：ＩＰ查看器<br>
价格：500 <b>金币</b><br>
功能：查看在线会员的IP</td>
<td width="15%" height="4" align="center" rowspan="2" bgcolor="FFFFFF"><FONT color=cccccc>
<IMG src="images/plus/colony.gif" border=1></FONT></td>
<td width="25%" height="2" bgcolor="FFFFFF">名称：讯息群发器<br>
价格：1000 <b>金币</b><br>
功能：群发讯息给在线会员</td>
</tr>
<tr>
<td width="25%" height="2" bgcolor="FFFFFF"><form method=POST action=?menu=lookip>
<input type="submit" value="我要购买" name="Submit"></td></form>
<td width="25%" height="2" bgcolor="FFFFFF">
<input type="submit" value="我要购买" name="Submit"  onclick="colony()"></td>
</tr>
</table>
</center>
</div>
</td>
</tr>



</table>




<%
rs.close
htmlend%>