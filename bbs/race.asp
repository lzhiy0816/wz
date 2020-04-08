<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>您还未<a href=login.asp>登录</a>社区")

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3

if Request.ServerVariables("request_method") = "POST" then


if Request("xuan")="" then
message=message&"<li>请选择马匹！"
end if

if not isnumeric(Request("money")) then
message=message&"<li>您的金币输入错误！"
end if

if Request("money") < 100 then
message=message&"<li>最低赌注：100 金币！"
end if

if rs("money") < int(Request("money")) then
message=message&"<li>您的金币不够！"
end if

if rs("userlife")<1 then
message=message&"<li>您没有体力了！"
end if



if message<>"" then
error(""&message&"")
end if


%>
<title>赛马实况</title>
<script>
document.oncontextmenu = new Function("return false;")
function Horse_timelinePlay(tmLnName, myID) {
if (document.Horse_Time == null) Horse_initTimelines();
tmLn = document.Horse_Time[tmLnName];
if(myID == null) { myID = ++tmLn.ID; firstTime=true;}
setTimeout('Horse_timelinePlay("'+tmLnName+'",'+myID+')',tmLn.delay);
fNew = ++tmLn.curFrame;
for (i=0; i<tmLn.length; i++) {
sprite = tmLn[i];
if (sprite.obj) {
numKeyFr = sprite.keyFrames.length; firstKeyFr = sprite.keyFrames[0];
if (fNew >= firstKeyFr && fNew <= sprite.keyFrames[numKeyFr-1]) {
keyFrm=1;
for (j=0; j<sprite.values.length; j++) {
props = sprite.values[j];
if (numKeyFr != props.length) {
if (props.prop2 == null) sprite.obj[props.prop] = props[fNew-firstKeyFr];
else sprite.obj[props.prop2][props.prop] = props[fNew-firstKeyFr];
}else{
while (keyFrm<numKeyFr && fNew>=sprite.keyFrames[keyFrm]) keyFrm++;
if (firstTime || fNew==sprite.keyFrames[keyFrm-1]) {
if (props.prop2 == null) sprite.obj[props.prop] = props[keyFrm-1];
else sprite.obj[props.prop2][props.prop] = props[keyFrm-1];
}}}}
}}
}






function Horse_initTimelines() {
document.Horse_Time = new Array();
document.Horse_Time[0] = new Array();
document.Horse_Time["horse"] = document.Horse_Time[0];

<%
Randomize



for i=0 to 8



for ii=1 to 29

d11=d11+(fix(rnd*40)+1)
tol=""&tol&","&d11&""
next




if d11 > Previous or i=1 then
best=i
Previous=d11
end if

%>








document.Horse_Time[0][<%=i%>]=new String("sprite");
document.Horse_Time[0][<%=i%>].obj=document.all ? document.all["h<%=i%>"] : null;
document.Horse_Time[0][<%=i%>].keyFrames=new Array(1,30);
document.Horse_Time[0][<%=i%>].values=new Array(2);
document.Horse_Time[0][<%=i%>].values[0]=new Array(0<%=tol%>);
document.Horse_Time[0][<%=i%>].values[0].prop="left";
document.Horse_Time[0][<%=i%>].values[1]=new Array();
document.Horse_Time[0][<%=i%>].values[0].prop2="style";

<%
d11=0
tol=""


next

if int(Request("xuan")) = best then
win="恭喜!你赢了 "&Request("money")*7&" 金币"
rs("money")=rs("money")+Request("money")*7
else
rs("money")=rs("money")-Request("money")
win="你输了 "&Request("money")&" 金币"
end if


rs("userlife")=rs("userlife")-1
rs.update
rs.close

%>


document.Horse_Time[0][9]=new String("sprite");
document.Horse_Time[0][9].obj = document.all ? document.all["sysMsg"] : null;
document.Horse_Time[0][9].keyFrames=new Array(1,30);
document.Horse_Time[0][9].values=new Array();
document.Horse_Time[0][9].values[0]=new Array();
document.Horse_Time[0][9].values[0].prop="left";
document.Horse_Time[0][9].values[1]=new Array();
document.Horse_Time[0][9].values[0]=new Array("hidden","inherit");
document.Horse_Time[0][9].values[0].prop="visibility";
document.Horse_Time[0][9].values[0].prop2="style";



for(i=0;i<document.Horse_Time.length;i++){document.Horse_Time[i].curFrame=1;document.Horse_Time[i].delay=200;}}




</script><body onLoad="Horse_timelinePlay('horse')">
<div id=h1 style=position:absolute;top:90px><img src=images/plus/h1.gif></div>
<div id=h2 style=position:absolute;top:130px><img src=images/plus/h2.gif></div>
<div id=h3 style=position:absolute;top:170px><img src=images/plus/h3.gif></div>
<div id=h4 style=position:absolute;top:210px><img src=images/plus/h4.gif></div>
<div id=h5 style=position:absolute;top:250px><img src=images/plus/h5.gif></div>
<div id=h6 style=position:absolute;top:290px><img src=images/plus/h6.gif></div>
<div id=h7 style=position:absolute;top:330px><img src=images/plus/h7.gif></div>
<div id=h8 style=position:absolute;top:370px><img src=images/plus/h8.gif></div>



<div id=sysMsg style=position:absolute;width:250;left:40%;><table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2>
<tr class=a1><td colspan=2 align=center>比赛结果公布  ---  <%=win%></td></tr><tr bgcolor=ffffff> <td width="40%" align=center>第１名：</td><td width="60%" align=center>(<%=best%>) <span id=ybb><%=best%></span></td> </tr></table></div>



<br><table width="100%" border="0" cellspacing="0" cellpadding="2" height="20"><tr>
<td><BR>您选择: <b><font color=red><%=Request("xuan")%></font> <span id=xuan><%=Request("xuan")%></span></b>　　赌注:<b><font color=red><%=Request("money")%></font> 金币</b></td>
<td align=right><BR><b>社区赛马场</b></td></tr></table>
<table width=100% border=0 cellspacing=1 cellpadding=5 class=a2>
<tr bgcolor=ffffff align=right><td height=39 width=10>１</td><td width=<%=Previous-15%>></td><td>艾基脑</td></tr>
<tr bgcolor=ffffff align=right><td height=39>２</td><td></td><td>马虎</td></tr>
<tr bgcolor=ffffff align=right><td height=39>３</td><td></td><td>赤白尾</td></tr>
<tr bgcolor=ffffff align=right><td height=39>４</td><td></td><td>雅玛侠</td></tr>
<tr bgcolor=ffffff align=right><td height=39>５</td><td></td><td>雷电</td></tr>
<tr bgcolor=ffffff align=right><td height=39>６</td><td></td><td>蒙面超人</td></tr>
<tr bgcolor=ffffff align=right><td height=39>７</td><td></td><td>MASAKI</td></tr>
<tr bgcolor=ffffff align=right><td height=39>８</td><td></td><td>极度凶兽</td>
</tr></table><center><br>　</body></html>



<A href="race.asp"><IMG src="images/plus/back.gif" 
border=0></A>


<script>
ybb.innerText=replace(ybb.innerText)
xuan.innerText=replace(xuan.innerText)

function replace(temp) {
temp = temp.replace(/1/ig,"艾基脑");
temp = temp.replace(/2/ig,"马虎");
temp = temp.replace(/3/ig,"赤白尾");
temp = temp.replace(/4/ig,"雅玛侠");
temp = temp.replace(/5/ig,"雷电");
temp = temp.replace(/6/ig,"蒙面超人");
temp = temp.replace(/7/ig,"MASAKI");
temp = temp.replace(/8/ig,"极度凶兽");
return (temp);
}
</script>


<%
htmlend
end if


%>
<title>社区 - 赛马中心</title><table width=100% border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td><img src="images/plus/h1.gif" width="36" height="35" align="absmiddle"> <font size="3"><b>赛马中心<br>
</b></font></td></tr></table><table border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td width="150" valign="top" align="center"><table width="90%" border="0" cellspacing="1" cellpadding="3" class=a2>
<tr class=a1><td><b>陪率:</b></td></tr><tr bgcolor=FFFFFF>
<td>
赔率 <font color=CC0000><b>1：7</b></font></td></tr><tr class=a1><td><b>状态</b></td>
</tr><tr><td bgcolor=FFFFFF><p>
体力 <b><%=rs("userlife")%></b><br>
金币 <b><%=rs("money")%></b> <font color="CC0000"><b>金币</b></font><br>
存款 <b><%=rs("savemoney")%></b> <font color="CC0000"><b>金币</b></font>
 </p></td></tr><tr class=a1>
<td><b>备注:</b></td></tr><tr><td bgcolor=FFFFFF align="center"><p>最小赌注 <b>100</b> <font color="CC0000"><b>金币</b></font><br>
最大赌注
<b>9999</b> <font color="CC0000"><b>金币</b></font>
<br>每次赌博体力将减 <FONT color=cc0000><B>1</B></FONT> 点

</td></tr></table></td><td width="10">　</td>
<td valign="top"><form METHOD="POST">
<table border=0 cellspacing="1" cellpadding="2" width="100%" class=a2><tr class=a1>
<td  align="center" width="29">闸位</td><td height="20"  align="center" width="45">马
图</td><td height="20"  align="center" width="54">马
名</td><td height="20"  align="center" width="44">出生地</td>
<td height="20"  align="center">简
介</td></tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>1</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h1.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>艾基脑</td><td width=44  align="center" bgcolor=FFFFFF>中国</td>
<td bgcolor=FFFFFF>AGENOW(艾基脑)有着惊人的坚强和冷静它的最大优势在于稳定的发挥和短长跑的多方面能力.</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>2</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h2.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>马虎</td><td width=44  align="center" bgcolor=FFFFFF>美国</td>
<td bgcolor=FFFFFF> MAHOO(马虎)其澎湃的速度和它的无处发泄的激情.在国外经受无数次大型赛事的磨练，经验丰富....<br>
</td></tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>3</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h3.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>赤白尾</td><td width=44  align="center" bgcolor=FFFFFF>台湾</td>
<td bgcolor=FFFFFF>这匹马的外观很奇特,白色的尾巴似乎在傲然宣告着自己不一般的血统.自出道以来就有着骄人的战绩...</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>4</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h4.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>雅玛侠</td><td width=44  align="center" bgcolor=FFFFFF>香港</td>
<td bgcolor=FFFFFF>体魄+斗志+意志+高贵.这一切它全部都拥有!均横的数值加上疯狂的耐力足以让它征服任何同类</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>5</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h5.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>雷电</td><td width=44  align="center" bgcolor=FFFFFF>英国</td>
<td bgcolor=FFFFFF>它神乎其神的身法和缥缈的步伐实在是让人叹为观止.速度神快，跑起来身上会闪出异常的光芒....</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>6</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h6.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>蒙面超人</td><td width=44  align="center" bgcolor=FFFFFF>印度</td>
<td bgcolor=FFFFFF><p>这匹马来历不明，暂时无资料...正在查...</p></td></tr><tr><td  align="center" width="29" bgcolor=FFFFFF><b>7</b></td>
<td  align="center" width="45" bgcolor=FFFFFF><img src="images/plus/h7.gif" width="36" height="35"></td>
<td  align="center" width="54" bgcolor=FFFFFF><font face="Arial, Helvetica, sans-serif"><b>MASAKI</b></font></td>
<td  align="center" width="44" bgcolor=FFFFFF>日本</td><td bgcolor=FFFFFF>来自日本打比大赛冠军，速度奇快,跑起来身上会像闪电一样,永往直前是它的本性.不受任何影响,直冲终点.</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>8</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h8.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>极度凶兽</td><td width=44  align="center" bgcolor=FFFFFF>南非</td>
<td bgcolor=FFFFFF>人不可貌相,马也一样.极度凶兽虽然其貌不扬，但是有一颗好斗的心.不服输的性格决定了它是不可限量的.</td>
</tr></table><div align="center"><table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr><td align="center"><p><select name="xuan" size="1"><option value=>□ 请选择马匹</option>
<option value=1>1道：艾基脑</option><option value=2>2道：马虎</option><option value=3>3道：赤白尾</option>
<option value=4>4道：雅玛侠</option><option value=5>5道：雷电</option><option value=6>6道：蒙面超人</option>
<option value=7>7道：MASAKI</option><option value=8>8道：极度凶兽</option></select>
赌注：<input type="text" name="money" value="100" size="4" maxlength="4"> <input type=submit value=开始比赛>
</td></tr></table>

<%
rs.close
htmlend%>