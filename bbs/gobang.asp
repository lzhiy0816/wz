<!-- #include file="setup.asp" -->
<%

top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

money=int(Request("money"))

if Request("menu")="win" then
if instr(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&""&Request.ServerVariables("script_name")&"") = 0 then error("<li>来源错误")
if Request.Cookies("gobang")=empty or Request.Cookies("gobang")>10 or Request.Cookies("sessionid")<>session.sessionid then error("<li>非法操作")
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
rs("money")=rs("money")+Request.Cookies("gobang")*2
rs.update
rs.close
Response.Cookies("gobang")=""
response.redirect "loading.htm"

elseif Request("menu")="lost" then
if Request.Cookies("gobang")=empty or Request.Cookies("gobang")>10 or Request.Cookies("sessionid")<>session.sessionid then error("<li>非法操作")
Response.Cookies("gobang")=""
response.redirect "loading.htm"
end if


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title>BBSxp五子棋</title>

<style>TABLE{BORDER-TOP:0px;BORDER-LEFT:0px;BORDER-BOTTOM:1px}TD{BORDER-RIGHT:0px;BORDER-TOP:0px}
</style>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td>　<img src="images/closedfold.gif" border="0">　<a href=main.asp><%=clubname%></a><br>
　<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">　<a href="Gobang.asp">五 子 棋</a></td>
</tr>
</table><br>
<center>
<%


if Request("menu")="" then
%><table cellSpacing="1" cellPadding="2" width="80%" border="0" class=a2>
	<tr class=a1>
		<td colspan="5" height="25" align="center"><strong>游戏介绍</strong></td>
	</tr>
	<tr class=a3>
		<td colspan="5"><p>五子棋是起源于中国古代的传统黑白棋种之一。现代五子棋日文称之为 “ 连珠 ” 
		，英译为 “Renju” ，英文称之为 “Gobang” 或 “FIR”(Five in a Row 的缩写 ) ，亦有 “ 连五子 ” 、 “ 
		五子连 ” 、 “ 串珠 ” 、 “ 五目 ” 、 “ 五目碰 ” 、 “ 五格 ” 等多种称谓。</p>
		<p>五子棋不仅能增强思维能力，提高智力，而且富含哲理，有助于修身养性。五子棋既有现代休闲的明显特征 “ 短、平、快 ” 
		，又有古典哲学的高深学问 “ 阴阳易理 ” 
		；它既有简单易学的特性，为人民群众所喜闻乐见，又有深奥的技巧和高水平的国际性比赛；它的棋文化源渊流长，具有东方的神秘和西方的直观；既有 “ 场 
		” 的概念，亦有 “ 点 ” 的连接。它是中西文化的交流点，是古今哲理的结晶。 </p>
		<p>五子棋起源于古代中国，发展于日本，风靡于欧洲。对于它与围棋的关系有两种说法，一说早于围棋，早在 “ 尧造围棋 ” 
		之前，民间就已有五子棋游戏；一说源于围棋，是围棋发展的一个分支。在中国的文化里，倍受人们的青睐。古代的五子棋的棋具与围棋相同，纵横各十七道。五子棋大约随围棋一起在我国南北朝时先后传入朝鲜、日本等地。据日本史料文献介绍，中国古代的五子棋是经由高丽 
		( 朝鲜 ) ，于 1688 年至 1704 年的日本元禄时代传到日本的。到日本明治 32 年 ( 公元 1899 年 ) ，经过公开征名， “ 
		连珠 ” 这一名称才被正式确定下来，取意于 “ 日月如合壁，五星如连珠 ” 。从此，连珠活动经过了不断的改良，主要是规则的变化 ( 
		即对执黑棋一方的限制 ) ，例如， 1899 年规定，禁止黑白双方走 “ 双三 ” ； 1903 年规定，只禁止黑方走 “ 双三 ” ； 
		1912 年规定，黑方被迫走 “ 双三 ” 亦算输； 1916 年规定，黑方不许走 “ 长连 ” ； 1918 年规定，黑方不许走 “ 
		四、三、三 ” ； 1931 年规定，黑方不许走 “ 双四 ” ，并规定将 19×19 的围棋盘改为 15×15 
		的连珠专用棋盘。本世纪初五子棋传入欧洲并迅速风靡全欧。通过一系列的变化，使五子棋这一简单的游戏复杂化、规范化，而最终成为今天的职业连珠五子棋，同时也成为一种国际比赛棋。 
		</p>
		</td>
	</tr>
	<tr class=a1>
		<td colspan="5" align="center" height="25">
		<p><strong>游戏规则 </strong></p>
		</td>
	</tr>
	<tr class=a3>
		<td colspan="5">1 、五子棋是玩家与电脑之间进行的竞技活动，玩家黑棋先行。<br>
		2 、五子棋专用盘为 15×15 ，五连子的方向为横、竖、斜。<br>
				3 、在棋盘上任何一方形成五连为胜。</td>
	</tr>
	<tr class=a4>
		<td width="30%" valign="bottom" align="center" height="25">
		<p><b>请选择您要下的赌注：</b></td>
		<td width=15% valign=bottom align="center" height="25"><b><a href=?menu=start&money=1>1金币</a></b></td>
		<td width="15%" valign="bottom" align="center" height="25"><b><a href="?menu=start&money=2">2金币</a></b></td>
		<td width="15%" valign="bottom" align="center" height="25"><b><a href="?menu=start&money=5">5金币</a></b></td>
		<td width="15%" valign="bottom" align="center" height="25"><b><a href="?menu=start&money=10">10金币</td>
	</tr>
	

	</table>

<%



else
if money > 10 or money < 1 then  error2("您的金币输入错误！")
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
if rs("money") < money then error2("您的金币不足！")
rs("money")=rs("money")-money
nowmoney=rs("money")
rs.update
rs.close
Response.Cookies("gobang")=money
Response.Cookies("sessionid")=session.sessionid

%>
<script>document.oncontextmenu = new Function("return false;")</script>

<FORM name=theform action=?menu=win method=post target=message></FORM>

<FORM name=lost action=?menu=lost method=post target=message></FORM>


<table border="0" cellspacing="0">

	<tr><td align="center">　</td><td align="center">
		您当前共有 <%=nowmoney%>  金币</td>
		<td align="center">本局赌注：<%=money%> 金币</td><td align="center">　</td></tr>
	<tr>
		<td align="center">
<img border="0" src="images/plus/Gobang1.gif"><br>
<br>
<img src="<%=userface%>"></td>



		<td background="images/plus/Gobangbg.gif" width="500" colspan="2">
		
		<div id="losshtml" style="position:absolute;visibility:hidden;"><table width=500 height=140 border=0 cellspacing=2 cellpadding=0><tr><td align=center><b><font size="7" color="#478130">五连珠，</font><font size="7" color="#FFFFFF">白</font><font size="7" color="#478130">胜!<br>
			</font><font size="3" color="#478130">很遗憾！您输了</font> <font size="3" color="#FF0000"><%=Request("money")%></font> <font size="3" color="#478130">金币.</font></b></td></tr></table></div>

		


<div id="winhtml" style="position:absolute;visibility:hidden;"><table width=500 height=140 border=0 cellspacing=2 cellpadding=0><tr><td align=center><b><font size="7" color="#478130">五连珠，</font><font size="7">黑</font><font size="7" color="#478130">胜!<br>
	</font><font color="#478130">恭喜您！您嬴了 </font><font size="3" color="#FF0000"><%=Request("money")%></font> <font size="3" color="#478130">金币.</font></b></td></tr></table></div>

		
		
<br>
<script src="inc/Gobang.js"></script>
</td>
		<td align="center">
<img border="0" src="images/plus/Gobang2.gif"><br>
<br><SCRIPT>
index = Math.floor(Math.random() * 84+1);
document.write("<img src=images/face/" + index + ".gif>");
</SCRIPT></td>
	</tr>
</table>




<%
end if
htmlend
%>