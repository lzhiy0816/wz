<!-- #include file="setup.asp" -->
<%
top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3

if Request.ServerVariables("request_method") = "POST" then


if Request("xuan")="" then
message=message&"<li>��ѡ����ƥ��"
end if

if not isnumeric(Request("money")) then
message=message&"<li>���Ľ���������"
end if

if Request("money") < 100 then
message=message&"<li>��Ͷ�ע��100 ��ң�"
end if

if rs("money") < int(Request("money")) then
message=message&"<li>���Ľ�Ҳ�����"
end if

if rs("userlife")<1 then
message=message&"<li>��û�������ˣ�"
end if



if message<>"" then
error(""&message&"")
end if


%>
<title>����ʵ��</title>
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
win="��ϲ!��Ӯ�� "&Request("money")*7&" ���"
rs("money")=rs("money")+Request("money")*7
else
rs("money")=rs("money")-Request("money")
win="������ "&Request("money")&" ���"
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
<tr class=a1><td colspan=2 align=center>�����������  ---  <%=win%></td></tr><tr bgcolor=ffffff> <td width="40%" align=center>�ڣ�����</td><td width="60%" align=center>(<%=best%>) <span id=ybb><%=best%></span></td> </tr></table></div>



<br><table width="100%" border="0" cellspacing="0" cellpadding="2" height="20"><tr>
<td><BR>��ѡ��: <b><font color=red><%=Request("xuan")%></font> <span id=xuan><%=Request("xuan")%></span></b>������ע:<b><font color=red><%=Request("money")%></font> ���</b></td>
<td align=right><BR><b>��������</b></td></tr></table>
<table width=100% border=0 cellspacing=1 cellpadding=5 class=a2>
<tr bgcolor=ffffff align=right><td height=39 width=10>��</td><td width=<%=Previous-15%>></td><td>������</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>��</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>���β</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>������</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>�׵�</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>���泬��</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>MASAKI</td></tr>
<tr bgcolor=ffffff align=right><td height=39>��</td><td></td><td>��������</td>
</tr></table><center><br>��</body></html>



<A href="race.asp"><IMG src="images/plus/back.gif" 
border=0></A>


<script>
ybb.innerText=replace(ybb.innerText)
xuan.innerText=replace(xuan.innerText)

function replace(temp) {
temp = temp.replace(/1/ig,"������");
temp = temp.replace(/2/ig,"��");
temp = temp.replace(/3/ig,"���β");
temp = temp.replace(/4/ig,"������");
temp = temp.replace(/5/ig,"�׵�");
temp = temp.replace(/6/ig,"���泬��");
temp = temp.replace(/7/ig,"MASAKI");
temp = temp.replace(/8/ig,"��������");
return (temp);
}
</script>


<%
htmlend
end if


%>
<title>���� - ��������</title><table width=100% border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td><img src="images/plus/h1.gif" width="36" height="35" align="absmiddle"> <font size="3"><b>��������<br>
</b></font></td></tr></table><table border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td width="150" valign="top" align="center"><table width="90%" border="0" cellspacing="1" cellpadding="3" class=a2>
<tr class=a1><td><b>����:</b></td></tr><tr bgcolor=FFFFFF>
<td>
���� <font color=CC0000><b>1��7</b></font></td></tr><tr class=a1><td><b>״̬</b></td>
</tr><tr><td bgcolor=FFFFFF><p>
���� <b><%=rs("userlife")%></b><br>
��� <b><%=rs("money")%></b> <font color="CC0000"><b>���</b></font><br>
��� <b><%=rs("savemoney")%></b> <font color="CC0000"><b>���</b></font>
 </p></td></tr><tr class=a1>
<td><b>��ע:</b></td></tr><tr><td bgcolor=FFFFFF align="center"><p>��С��ע <b>100</b> <font color="CC0000"><b>���</b></font><br>
����ע
<b>9999</b> <font color="CC0000"><b>���</b></font>
<br>ÿ�ζĲ��������� <FONT color=cc0000><B>1</B></FONT> ��

</td></tr></table></td><td width="10">��</td>
<td valign="top"><form METHOD="POST">
<table border=0 cellspacing="1" cellpadding="2" width="100%" class=a2><tr class=a1>
<td  align="center" width="29">բλ</td><td height="20"  align="center" width="45">��
ͼ</td><td height="20"  align="center" width="54">��
��</td><td height="20"  align="center" width="44">������</td>
<td height="20"  align="center">��
��</td></tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>1</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h1.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>������</td><td width=44  align="center" bgcolor=FFFFFF>�й�</td>
<td bgcolor=FFFFFF>AGENOW(������)���ž��˵ļ�ǿ���侲����������������ȶ��ķ��ӺͶ̳��ܵĶ෽������.</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>2</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h2.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>��</td><td width=44  align="center" bgcolor=FFFFFF>����</td>
<td bgcolor=FFFFFF> MAHOO(��)�����ȵ��ٶȺ������޴���й�ļ���.�ڹ��⾭�������δ������µ�ĥ��������ḻ....<br>
</td></tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>3</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h3.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>���β</td><td width=44  align="center" bgcolor=FFFFFF>̨��</td>
<td bgcolor=FFFFFF>��ƥ�����ۺ�����,��ɫ��β���ƺ��ڰ�Ȼ�������Լ���һ���Ѫͳ.�Գ������������Ž��˵�ս��...</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>4</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h4.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>������</td><td width=44  align="center" bgcolor=FFFFFF>���</td>
<td bgcolor=FFFFFF>����+��־+��־+�߹�.��һ����ȫ����ӵ��!�������ֵ���Ϸ��������������������κ�ͬ��</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>5</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h5.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>�׵�</td><td width=44  align="center" bgcolor=FFFFFF>Ӣ��</td>
<td bgcolor=FFFFFF>����������������翵Ĳ���ʵ��������̾Ϊ��ֹ.�ٶ���죬���������ϻ������쳣�Ĺ�â....</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>6</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h6.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>���泬��</td><td width=44  align="center" bgcolor=FFFFFF>ӡ��</td>
<td bgcolor=FFFFFF><p>��ƥ��������������ʱ������...���ڲ�...</p></td></tr><tr><td  align="center" width="29" bgcolor=FFFFFF><b>7</b></td>
<td  align="center" width="45" bgcolor=FFFFFF><img src="images/plus/h7.gif" width="36" height="35"></td>
<td  align="center" width="54" bgcolor=FFFFFF><font face="Arial, Helvetica, sans-serif"><b>MASAKI</b></font></td>
<td  align="center" width="44" bgcolor=FFFFFF>�ձ�</td><td bgcolor=FFFFFF>�����ձ���ȴ����ھ����ٶ����,���������ϻ�������һ��,����ֱǰ�����ı���.�����κ�Ӱ��,ֱ���յ�.</td>
</tr><tr><td width=29  align="center" bgcolor=FFFFFF><b>8</b></td><td width=45  align="center" bgcolor=FFFFFF><img src="images/plus/h8.gif" width="36" height="35"></td>
<td width=54  align="center" bgcolor=FFFFFF>��������</td><td width=44  align="center" bgcolor=FFFFFF>�Ϸ�</td>
<td bgcolor=FFFFFF>�˲���ò��,��Ҳһ��.����������Ȼ��ò���������һ�źö�����.��������Ը���������ǲ���������.</td>
</tr></table><div align="center"><table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr><td align="center"><p><select name="xuan" size="1"><option value=>�� ��ѡ����ƥ</option>
<option value=1>1����������</option><option value=2>2������</option><option value=3>3�������β</option>
<option value=4>4����������</option><option value=5>5�����׵�</option><option value=6>6�������泬��</option>
<option value=7>7����MASAKI</option><option value=8>8������������</option></select>
��ע��<input type="text" name="money" value="100" size="4" maxlength="4"> <input type=submit value=��ʼ����>
</td></tr></table>

<%
rs.close
htmlend%>