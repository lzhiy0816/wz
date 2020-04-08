<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>

<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img src="images/coner.gif"><img src="images/openfold.gif">　<a href="play.asp">社区赌场</a></font></td>
</tr>
</table><br>
<%



select case Request("menu")
case ""
play

case "dx"
dx

case "ds"
ds

case "sz"
sz

case "dx1"
checkup
dx1

case "ds1"
checkup
ds1

case "sz1"
checkup
sz1

end select

sub checkup
money=int(request("money"))
if money > rs("money") then
error2("你的金币不够！")
end if
if money<10 or money>100 then
error2("输入错误！赌注范围10-100")
end if

if rs("userlife")<1 then
error2("您没有体力了!")
end if
end sub



sub play%>
<title>社区博彩</title><BODY><br>

<center>
<table width="600" border="0" height="143" cellpadding="1" cellspacing="1" class=a2>
<tr align="center">
<td height="129">
<table border="0" width="100%" cellspacing="0" cellpadding="6"><tr>
<td width="50%" align="center" bgcolor="FFFFFF">警告：每次赌博，您的魅力值会扣 1 ！
<p><img src="images/plus/daxiao.gif" width="303" height="237" border="0"></td>
<td width="50%" class=a4>
<ul><li><b><a href="play.asp?menu=sz">赌骰子</a></b></font><p>游戏规则：庄家随机出三个<b>骰子</b>，玩家也随机出三个<b>骰子</b>，谁大，谁赢，赔率１：１</font></p>
<blockquote><p align="right"><a href="play.asp?menu=sz">现在进入&gt;&gt;</a></p></blockquote>



<li><b><a href="play.asp?menu=ds">猜点数</a></b><p>游戏规则：电脑随机出一个<b>骰子</b>，您选择骰子的点数，如果<b>骰子</b>的点数和您选的点数一样，您就赢了，赔率１：５</font></p>
<ul>
<blockquote><p align="right"><a href="play.asp?menu=ds">现在进入&gt;&gt;</a></p></blockquote>
</li></ul>

<li><b><a href="play.asp?menu=dx">赌围骰</a></b><p>游戏规则：电脑随机出三个<b>骰子</b>，您选择可能会出现的</font>三个<b>骰子</b>，如果<b>骰子</b>和您选的一样，您就赢了，赔率１：１００</font></p>
</li></ul>
<blockquote><p align="right"><a href="play.asp?menu=dx">现在进入&gt;&gt;</a></p></blockquote>

</td></tr></table></td></tr></table></center>

<%end sub

sub dx%>
<br>

<form method="POST" action="play.asp?menu=dx1"><input type="hidden" name=menu value=dxover>

<center>
<table class=a2 cellSpacing="1" width="400" border="0" height="54" cellpadding="6">
<tr align="center" class=a1><td width="190" colspan="2">
<b>
赌 场 - 赌 围 骰</b></font></td>
<td width="183" colspan="2"><p>最大下注纹银是 <b><font color="CC0000">
100</font> 金币</b></p>
</td>
</tr>
<tr align="center"><td width="134" bgcolor="FFFFFF"><img src="images/plus/run.gif" width="38" height="36"></td>
<td width="109" colspan="2" bgcolor="FFFFFF"><img src="images/plus/run.gif" width="38" height="36"></td>
<td width="117" bgcolor="FFFFFF"><img src="images/plus/run.gif" width="38" height="36"></td>
</tr><tr><td class=a1 width="386" colspan="4">
<p align="center"><font style=font-size:9pt><b>
　请选择</b></font></td>
</tr><tr><td align=center colspan="4" width="386" bgcolor="FFFFFF" >
<table border="0" cellspacing="0" cellpadding="0" width="386">
<tr align="center"><td width="65">
<p align="center"><img src="images/plus/w1.gif"></td>
<td width="65">
<img src="images/plus/w2.gif"></td>
<td width="64">
<img src="images/plus/w3.gif"></td>
<td width="64">
<img src="images/plus/w4.gif"></td>
<td width="64">
<img src="images/plus/w5.gif"></td>
<td width="64">
<img src="images/plus/w6.gif"></td>
</tr>
<tr align="center"><td width="65">
<input type="radio" name="dddimg" value="1" checked></td>
<td width="65">
<input type="radio" name="dddimg" value="2"></td>
<td width="64">
<input type="radio" name="dddimg" value="3"></td>
<td width="64">
<input type="radio" name="dddimg" value="4"></td>
<td width="64">
<input type="radio" name="dddimg" value="5"></td>
<td width="64">
<input type="radio" name="dddimg" value="6"></td>
</tr></table></td></tr><tr>
<td align=center colspan="4" width="386" class=a3>我要下注：
<input type="text" name="money" size="10" value="10"> <b> 金币</b></td></tr><tr>
<td align=center colspan="4" width="386" bgcolor="FFFFFF"><input type="submit" value="下注啦！！！" name="B1" style="font-size: 12px">
</td>
</tr></table></center>
</form><center>
警告：每次赌博，您的体力值会扣 <font color="CC0000">1</font> ！
<br>您现在一共有社区纹银  <b><font color="CC0000"><%=rs("money")%></font></b> <b>金币</b> 可以作为赌注<br>

<%end sub

sub ds%>
<br>
<form method="POST" action="play.asp?menu=ds1"><input type="hidden" name=menu value=dsover>
<center>
<table border=0 cellspacing=1 cellpadding=3 width="350" class=a2>
<tr align="center" class=a1><td width="33%"><p align="center">
<b>赌 场 - 猜 点 数</b></font></p>
</td></tr>
<tr align="center"><td width="33%" bgcolor="FFFFFF"><img src="images/plus/run.gif"></td></tr><tr>
<td class=a1 width="960" align="center"><b>请选择</b></font></td>
</tr><tr><td align=center bgcolor="FFFFFF"><table border="0" cellspacing="0" cellpadding="0"><tr align="center">
<td width="17%"><img src="images/plus/dice1.gif"></td><td width="17%"><img src="images/plus/dice2.gif"></td>
<td width="16%"><img src="images/plus/dice3.gif"></td><td width="17%"><img src="images/plus/dice4.gif"></td>
<td width="17%"><img src="images/plus/dice5.gif"></td><td width="16%"><img src="images/plus/dice6.gif"></td>
</tr><tr align="center"><td width="17%"><input type="radio" name="dddimg" value="1" checked>
</td><td width="17%"><input type="radio" name="dddimg" value="2"></td><td width="16%">
<input type="radio" name="dddimg" value="3"></td><td width="17%"><input type="radio" name="dddimg" value="4">
</td><td width="17%"><input type="radio" name="dddimg" value="5"></td><td width="16%">
<input type="radio" name="dddimg" value="6"></td></tr></table></td></tr><tr><td align=center class=a3>
我要下注：
<input type="text" name="money" size="10" value="10"> <b> 金币</b></td></tr><tr>
  <td align=center bgColor=#FFFFFF><input type="submit" value="下注啦！！！" name="B1" style="font-size: 12px">
</td>
</tr></table>  </form>
警告：每次赌博，您的体力值会扣 <font color="CC0000">1</font> ！<br>您现在一共有社区纹银 <b><font color="CC0000"><%=rs("money")%></font> 金币</b> 可以作为赌注
<br>
<%end sub

sub sz%>
<form method="POST" action="play.asp?menu=sz1"><input type="hidden" name=menu value=szover>
<center>
<table border=0 cellspacing=1 cellpadding=6 width="400" class=a2 height="115">
<tr class=a1>
<td width="160" class=a2 height="9" align="center">
<b>赌场 - 赌骰子</b></font></td>
<td width="160" class=a2 height="9" align="right">最大下注纹银是 <b><font color="CC0000">100</font> 金币</b></td>
</tr><tr>
<td align=center class=a3 height="12" colspan="2">您现在一共有社区纹银 <b><font color="CC0000"><%=rs("money")%></font>
金币</b> 可以作为赌注</td></tr><tr>
<td align=center class=a4 height="20" colspan="2">我要下注：
<input type="text" name="money" size="10" value="10"> <b> 金币</b></td></tr><tr>
<td align=center bgColor=ffffff height="23" colspan="2"><input type="submit" value="下注啦！！！" name="B1" style="font-size: 12px"></td>
</tr></table>
</form>
警告：每次赌博，您的体力值会扣 <font color="CC0000">1</font> ！<br>

<%end sub



sub dx1
money=request("money")
dim ds
ds=request("dddimg")
Randomize
d1=int(rnd*6)+1
d2=int(rnd*6)+1
d3=int(rnd*6)+1
%>
<center><font style=font-size:9pt><b><font color="0000CC">您押的是</font><br><font color="CC0000"></font><font color="CC0000">
</font></b></font><img src=images/plus/w<%=ds%>.gif><font style=font-size:9pt><b><font color="CC0000"></font></b></font>
<center>
<table class=a2 cellSpacing="1" width="400" border="0" height="54" cellpadding="6">
<tr><td colspan="3" align="center" class=a1><b>结果:</b></font></td>
</tr><tr class=a3><td width="33%" align="center"><img src=images/plus/dice<%=d1%>.gif></td>
<td width="33%" align="center"><img src=images/plus/dice<%=d2%>.gif></td>
<td width="34%" align="center"><img src=images/plus/dice<%=d3%>.gif></td></tr>
<tr><td colspan="3" align="center" class=a3> <br>
<%
if ds=""&d1&"" and ds=""&d2&"" and ds=""&d3&"" then
%>
<p><font size="3">嘻嘻~，赢了： <b><font color="CC0000"><%=money*100%></font> 金币</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="000099">庄家：</font></b></td>
<td colspan="2">您真厉害呀~！还要来吗？</td></tr><tr><td align="right"><b><font color="CC0000">我说：</font></b></td>
<td colspan="2"><a href="play.asp?menu=dx">再来，我今天运气不错耶~！</a></td></tr>
<tr><td align="right">　</td><td colspan="2"><a href=main.asp>见好就收，呵，您都傻的~！</a></td>
</tr></table>
<%
money=rs("money")+money*100
else

%>
<p><font size="3">真倒霉，输了： <b><font color="CC0000"><%=money%></font> 金币</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="000099">庄家：</font></b></td>
<td colspan="2">呵呵还要来吗？ 有赌未为输哦~！</td></tr><tr><td align="right"><b><font color="CC0000">我说：</font></b></td>
<td colspan="2"><a href="play.asp?menu=dx">再来，呸，我就不相信这么倒霉~！</a></td>
</tr><tr><td align="right">　</td><td colspan="2"><a href=main.asp>今天真倒霉，不来了，下次再玩~！</a></td>
</tr></table>
<%
money=rs("money")-money
end if
rs("money")=money
rs("userlife")=rs("userlife")-1
rs.update
%>
<hr width="250" SIZE="1">
您现在有社区金币：<b><%=rs("money")%>&nbsp; </b>体力:<%=rs("userlife")%>
<hr width="250" SIZE="1"></td></tr></table></center><font size="3"></font>
<%end sub



sub ds1
money=request("money")
dim ds
ds=cint(request("dddimg"))
Randomize
d1=fix(rnd*6)+1
%>
<br>
<center>
<table border="0" cellpadding="3" width="400" cellspacing="1" class=a2>
<tr class=a1><td align="center"><b>您押的是</b></td>
<td align="center"><b>结果:</b></font><b><%=d1%>点</b></td>
</tr><tr><td align="center" class=a3><img src="images/plus/dice<%=ds%>.gif"></td>
<td align="center" class=a3><Img src="images/plus/dice<%=d1%>.gif">
</span></a></td></tr>

<%
if ds=d1 then
money=rs("money")+money*5
%>
<tr><td colspan="2" align="center"  class=a3 width="392"> <br><p><font size="3">嘻嘻~，赢了： <b><font color="CC0000"><%=request("money")*5%></font> 金币</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="000099">庄家：</font></b></td>
<td colspan="2">您真厉害呀~！还要来吗？</td></tr><tr><td align="right"><b><font color="CC0000">我说：</font></b></td>
<td colspan="2"><a href="play.asp?menu=ds">再来，我今天运气不错耶~！</a></td></tr>
<tr><td align="right">　</td><td colspan="2"><a href=main.asp>见好就收，呵，您都傻的~！</a></td>
</tr></table>
<%else%>
<tr><td colspan="2" align="center" class=a3 width="392"> <br><p><font size="3">真倒霉，输了： <b><font color="CC0000"><%=money%> </font> 金币</b></font></p>
<table width="80%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right"><b><font color="000099">庄家：</font></b></td>
<td colspan="2">呵呵还要来吗？ 有赌未为输哦~！</td></tr><tr><td align="right"><b><font color="CC0000">我说：</font></b></td>
<td colspan="2"><a href=play.asp?menu=ds>再来，呸，我就不相信这么倒霉~！</a></td>
</tr><tr><td align="right">　</td><td colspan="2"><a href=main.asp>今天真倒霉，不来了，下次再玩~！</a></td>
</tr></table>
<%
money=rs("money")-money
end if
rs("money")=money
rs("userlife")=rs("userlife")-1
rs.update
%>
<hr width="250" SIZE="1">
您现在有社区金币：<b><font color="CC0000"><%=rs("money")%></font>&nbsp; </b>体力:<%=rs("userlife")%>
<hr width="250" SIZE="1"></td></tr></table></center>
<font size="3"></font>
<%end sub

sub sz1
money=request("money")
Randomize
d1=fix(rnd*6)+1
d2=fix(rnd*6)+1
d3=fix(rnd*6)+1
b1=fix(rnd*6)+1
b2=fix(rnd*6)+1
b3=fix(rnd*6)+1
zj=d1+d2+d3
ni=b1+b2+b3
%>
<br>
<center>
<table border="0" cellspacing="1" cellpadding="6" width="400" class=a2 height="314">
<tr class=a1><td colspan="3" height="12" align="center"><b>赌 场 - 赌 骰 子</b></font></td></tr><tr>
<td colspan="3" align="center" class=a4 height="12"><font color="0000CC">□-庄家骰子：</font><%=zj%>
点 </td></tr><tr><td align="center" height="30" bgcolor="FFFFFF"><img src=images/plus/dice<%=d1%>.gif></td>
<td align="center" height="30" bgcolor="FFFFFF"><img src=images/plus/dice<%=d2%>.gif></td>
<td align="center" height="30" bgcolor="FFFFFF"><img src=images/plus/dice<%=d3%>.gif></td>
</tr><tr><td colspan="3" align="center" class=a4 height="12"><font color="CC0000">□-您的骰子：</font><%=ni%>
点 </td></tr><tr><td align="center" height="30" bgcolor="FFFFFF"><img src=images/plus/dice<%=b1%>.gif></td>
<td align="center" height="30" bgcolor="FFFFFF"><img src=images/plus/dice<%=b2%>.gif></td>
<td align="center" height="30" bgcolor="FFFFFF"><img src=images/plus/dice<%=b3%>.gif></td>
</tr><tr>
<td colspan="3" align="center" class=a4 height="116">
<%
if zj<ni then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="5"><tr><td align="right" height="50">　</td>
<td colspan="2" height="50"><font size="3">&nbsp;嘻嘻~，赢了：<b><font color="CC0000"><%=money%></font> 金币</b></font></td></tr><tr>
<td align="right"><b><font color="000099"> 庄家：</font></b></td><td colspan="2">&nbsp;您真厉害呀~！还要来吗？ </td></tr><tr><td align="right"><b><font color="CC0000">我要：</font></b></td>
<td colspan="2">&nbsp;<a href="play.asp?menu=sz">再来，我今天运气不错耶~！</a></td></tr><tr><td align="right">　</td>
<td colspan="2">&nbsp;<a href=main.asp>见好就收，呵呵，走人咯~！</a></td></tr></table>
<%
money=rs("money")+money
else
%>
<center>
<table width="100%" border="0" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="111111"><tr>
<td height="50">　</td><td colspan="2" height="50"><p><font size="3">&nbsp;真倒霉，输了： <b><font color="CC0000"><%=money%></font> 金币</b></font></p></td></tr><tr>
<td align="right"><b><font color="000099"> 庄家：</font></b></td><td colspan="2">&nbsp;呵呵~还要来吗？ 有赌未必输哦~！</td></tr><tr><td align="right" rowspan="2"><b><font color="CC0000"> 我要：</font></b></td>
<td colspan="2">&nbsp;<a href="play.asp?menu=sz">再来，呸，我就不相信这么倒霉~！</a></td></tr><tr><td colspan="2">&nbsp;<a href=main.asp>今天真倒霉，走~了，下次再玩~！</a></td></tr></table>
</center>
<%
money=rs("money")-money
end if

rs("money")=money
rs("userlife")=rs("userlife")-1
rs.update
%>
<hr size="1" width="250">
您现在有社区金币：<font color="CC0000"><b><%=rs("money")%></b></font><b> </b>&nbsp;体力：<%=rs("userlife")%>
<hr size="1" width="250"></td></tr></table></center>
<%end sub
rs.close
htmlend
%>