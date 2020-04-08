<!-- #include file="setup.asp" -->
<!-- #include file="inc/MD5.asp" -->

<%

if CloseRegUser = 1 then
error("<li>本论坛暂时不开放新用户注册！")
end if

username=server.htmlencode(Trim(Request("username")))

if Request("menu")="Check" then

If conn.Execute("Select id From [user] where username='"&username&"'" ).eof Then
response.write "用户名&quot; <font color=red>"&HTMLEncode(username)&"</font> &quot;可以正常注册！"
else
response.write "您所选的用户名&quot; <font color=red>"&username&"</font> &quot;已经有用户使用，请另外选择一个用户名。"
end if

responseend

end if





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.ServerVariables("request_method") = "POST" then
password=Trim(Request("password"))
userpass2=Trim(Request("userpass2"))
usermail=server.htmlencode(Request("usermail"))
userhome=server.htmlencode(Request("userhome"))
question=server.htmlencode(Request("question"))
answer=server.htmlencode(Request("answer"))
realname=server.htmlencode(Request("realname"))
birthday=server.htmlencode(Request("birthday"))
userface=server.htmlencode(Request("userface"))
sex=server.htmlencode(Request("sex"))
country=server.htmlencode(Request("country"))
province=server.htmlencode(Request("province"))
city=server.htmlencode(Request("city"))
blood=server.htmlencode(Request("blood"))
character=server.htmlencode(Request("character"))
belief=server.htmlencode(Request("belief"))
college=server.htmlencode(Request("college"))
marital=server.htmlencode(Request("marital"))
education=server.htmlencode(Request("education"))
personal=""&server.htmlencode(Request("personal"))&""
occupation=server.htmlencode(Request("occupation"))
userqq=server.htmlencode(Request("userqq"))
icq=server.htmlencode(Request("icq"))
sign=server.htmlencode(Request("sign"))
sign=replace(sign,vbCrlf,"<br>")
sign=replace(sign,"\","&#92;")



temp=UCase(sign)
if instr(temp,"[FLASH]")>0 or instr(temp,"[RM]")>0 or instr(temp,"[MP]")>0 then
message=message&"<li>签名档中不能含有[FLASH] [RM] [MP]代码"
end if


if instr(username,"	")>0 or instr(username,"#")>0 or instr(username,"`")>0 or instr(username,"|")>0 or instr(username," ")>0 or instr(username,"　")>0 or Instr(username,"%")>0 or Instr(username,"&")>0 or Instr(username,"ヴ")>0 or Instr(username,"ヂ")>0 or Instr(username,"ゼ")>0 or Instr(username,"ヅ")>0 or Instr(username,"") or Instr(username,"+")>0 then
message=message&"<li>您的用户名中不能含有特殊符号"
end if


if username="" then
message=message&"<li>您的用户名没有填写"
end if

if Len(username)>16 then
message=message&"<li>您的用户名中不能超过16个字节"
end if

if password="" then
password=session.sessionid
elseif password<>userpass2 then
message=message&"<li>您2次输入的密码不正确"
end if

if instr(usermail,"@")=0 then
message=message&"<li>您的电子邮件地址填写错误"
end if


if not isnumeric(userqq) and userqq<>"" then
message=message&"<li>QQ号码请用数字填写"
end if

if not isnumeric(icq) and icq<>"" then
message=message&"<li>ICQ号码请用数字填写"
end if


if Len(personal)>255 then
message=message&"<li>个人简介不能大于 255 个字节"
end if

if Len(sign)>255 then
message=message&"<li>签名档不能大于 255 个字节"
end if


sql="select * from [user] where username='"&HTMLEncode(username)&"'"
rs.Open sql,Conn
if not rs.eof then
message=message&"<li>此用户名已经被别人注册了"
end if
rs.close



if RegOnlyMail = 1 then
sql="select * from [user] where usermail='"&usermail&"'"
rs.Open sql,Conn
if not rs.eof then
message=message&"<li>此Email已经被别人注册了"
end if
rs.close
end if






if message<>"" then
error(""&message&"")
end if


for each ho in request.form("character")
allcharacter=""&allcharacter&""&ho&""
next


rs.Open "[user]",conn,1,3

rs.addnew
rs("username")=username
rs("userpass")=md5(password)
rs("usermail")=usermail
rs("userhome")=userhome
rs("question")=question
if Request("answer")<>empty then rs("answer")=md5(Request("answer"))
rs("realname")=realname
rs("birthday")=birthday
rs("userface")=userface
rs("sex")=sex
rs("country")=country
rs("province")=province
rs("city")=city
rs("blood")=blood
rs("character")=allcharacter
rs("belief")=belief
rs("occupation")=occupation
rs("college")=college
rs("marital")=marital
rs("education")=education
if icq<>empty then rs("icq")=icq
if userqq<>empty then rs("userqq")=userqq
rs("personal")=personal
rs("sign")=sign
rs("friend")="|"
rs("regtime")=""&Date()&""
rs.update
rs.close


Application("NewUserName")=username


mailaddress=usermail
mailtopic="用户名注册成功"
body=""&vbCrlf&"亲爱的"&username&", 您好!"&vbCrlf&""&vbCrlf&"　　恭喜! 您已经成功地注册了您的资料, 非常感谢您使用"&homename&"的服务!"&vbCrlf&""&vbCrlf&"　* 您的帐号是:"&username&"　密码是:"&password&""&vbCrlf&""&vbCrlf&"　* "&clubname&"("&cluburl&"/Default.asp)"&vbCrlf&""&vbCrlf&"　* 最后, 有几点注意事项请您牢记"&vbCrlf&"1、请遵守《计算机信息网络国际联网安全保护管理办法》里的一切规定。"&vbCrlf&"2、使用轻松而健康的话题，所以请不要涉及政治、宗教等敏感话题。"&vbCrlf&"3、承担一切因您的行为而直接或间接导致的民事或刑事法律责任。"&vbCrlf&""&vbCrlf&""&vbCrlf&"论坛服务由 "&homename&"("&homeurl&") 提供　程序制作：Yuzi工作室(http://www.yuzi.net)"&vbCrlf&""&vbCrlf&""&vbCrlf&""
%>
<!-- #include file="inc/mail.asp" -->
<%



if SendPassword<>1 then
userword="<li>用户名：<font color=red>"&username&"</font><li>密码：<font color=red>"&password&"</font>"
Response.Cookies("username")=username
Response.Cookies("userpass")=md5(password)
end if


message=message&""&userword&"<li>注册新用户资料成功<li><a href=main.asp>返回论坛首页</a>"
succeed(""&message&"<meta http-equiv=refresh content=3;url=main.asp>")



end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>


<title>注册用户资料</title>

<script>function New(para_URL){var URL=new String(para_URL);window.open(URL,'','resizable,scrollbars')}
function Check(){var Name=document.form.username.value;
window.open("register.asp?menu=Check&username="+Name,"","width=200,height=20");
}
</script>
<SCRIPT src="inc/birthday.js"></SCRIPT>



<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
<td vAlign="center" align="top">&nbsp;<font color=000000><img src="images/closedfold.gif" width="14" height="14">　<a href=main.asp><%=clubname%></a><br>
&nbsp;<img height=15 src="images/coner.gif" width=15><img src="images/openfold.gif">　注册新用户</font></td>
</tr>
</table>






<br>

<SCRIPT>valigntop()</SCRIPT>
<table width=97% align=center cellspacing=0 cellpadding=0 border=0 class=a1> <form method="POST" name="form" onsubmit="return VerifyInput();">
<tr>
<td>
<table width=100% cellspacing=1 cellpadding=4 border=0 class=a2>
<tr>
<td height="20" align="center" colspan="2" class=a1><b>填写用户资料</b></td>
</tr>
<tr bgcolor="FFFFFF">
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;<img src="images/1.gif">个人社区信息</b>（以下内容必填）</td>
</tr>
<tr>
<td class=a3 height="5" align="right" valign="middle" width="46%"><b>用户名：</b></td>
<td class=a3 height="5" align="left" valign="middle" width="54%">
&nbsp;
<input type="text" name="username" size="28" maxlength="12" value="<%=Request("username")%>">&nbsp;<input onclick="javascript:Check()" type="button" value="检测帐号" name="button">
</td>
</tr>
<%if SendPassword=0 then%>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>密码：</b></td>
<td class=a4 height="2" align="left" valign="middle" width="54%">
&nbsp;
<input type="password" name="password" size="40" maxlength="16">
</td>
</tr>
<tr class=a3>
<td height="2" align="right" valign="middle" width="46%"><b>确认密码：</b><br>请与您的密码保持一致</td>
<td height="2" align="left" valign="middle" width="54%" class=a3>
&nbsp;
<input type="password" name="userpass2" size="40">
</td>
</tr>
<%end if%>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>您的Email地址：</b><br>
<font color="FF0000">密码将通过Email发送</font></td>
<td class=a4 height="2" align="left" valign="middle" width="54%">
&nbsp;
<input type="text" name="usermail" size="40">
</td>
</tr>
<tr bgcolor="FFFFFF">
<td height="25" align="left" valign="middle" colspan="2"><b>&nbsp;<img src="images/1.gif">个人生活信息</b>（以下内容建议填写）</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a3 width="46%"> 　<b>真实姓名：</b>
<input type="text" name="realname" size="18">
</td>
<td height="71" align="left" valign="top" class=a4 rowspan="14" width="54%">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
<tr>
<td><b>头<font color="000000">　</font>像：</b> 


<script>function showimage(){document.images.tus.src=""+document.form.userface.options[document.form.userface.selectedIndex].value+"";}

facetol=84;
temp = Math.floor(Math.random() * facetol+1);


document.write("<img src=images/face/"+temp+".gif name=tus> ")
document.write("<select name=userface size=1 onChange=showimage()>")

for(i=1;i<=facetol;i=i+1) {
if (i==temp) ii="selected";else ii=""
document.write("<option value='images/face/"+i+".gif' "+ii+">"+i+"</option>")
}
</script>


</select> <a href="JavaScript:New('face.htm')">查看所有的头像列表</a></td>
</tr>
<tr>
<td><b>生　日：</b> <input onfocus="show_cele_date(birthday,'','',birthday)" value="" name="birthday"></td>
</tr>
<tr>
<td><b><font color="000000">性　格</font>： &nbsp; </b>
<br>
<SCRIPT>
var moderated="成熟稳重 幼稚调皮 温柔体贴 活泼可爱 脾气暴躁 内向害羞 外向开朗 心地善良 风趣幽默 慢条斯理 积极进取 郁郁寡欢 处事洒脱 圆滑老练 豪放不羁 异想天开 多愁善感 淡泊名利 情绪多变 胆小怕事 循规蹈矩 热心助人 快言快语 爱管闲事 追求刺激"
var list= moderated.split (' '); 
for(i=0;i<list.length;i=i+1) {
if (i==5 || i==10 || i==15 || i==20) document.write("<br>")
if (list[i] !=""){document.write("　 <input type=checkbox value=' "+list[i]+"' name=character id=character"+i+"> <label for=character"+i+">"+list[i]+"</label>")}
}
</SCRIPT>

</td>
</tr>
<tr>
<td><b>个人简介： &nbsp;
<textarea name=personal rows=6 cols=65></textarea>
</b></td>
</tr>
<tr>
<td height="10"><b>签名档： &nbsp;
<textarea name=sign rows=6 cols=65></textarea> </b> </td>
</tr>
</table>
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%"><b> 　性　　别：</b>
<select size=1 name=sex>
<option value="" selected>[请选择]</option>
<option value=male>男</option>
<option value=female>女</option>
</select>
</td>
</tr>

<tr class=a3>
<td height="2" valign="top" width="46%">　<b>国　　家：</b>
<b>
<input type="text" name="country" size="18">
</b> </td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">　<b>省　　份：</b>
<input type="text" name="province" size="18">
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">　<b>城　　市：
</b>
<input type="text" name="city" size="18">
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">
　<b>血　　型：</b>
<select size=1 name=blood>
<option value="" selected>[请选择]</option>
<option
value=A>A</option>
<option value=B>B</option>
<option
value=AB>AB</option>
<option value=O>O</option>
<option
value=其他>其他</option>
</select>
</td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">　<b>信　　仰：</b>
<select size=1 name=belief>
<option value="" selected>[请选择]</option>
<option value=佛教>佛教</option>
<option
value=道教>道教</option>
<option value=基督教>基督教</option>
<option
value=天主教>天主教</option>
<option value=回教>回教</option>
<option
value=无神论者>无神论者</option>
<option value=共产主义者>共产主义者</option>
<option
value=其他>其他</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">　<b>职　　业： </b>
<select name="occupation">
<option value="" selected>[请选择]</option>
<option value="财会/金融">财会/金融</option>
<option value="工程师">工程师</option>
<option value="顾问">顾问</option>
<option value="计算机相关行业">计算机相关行业</option>
<option value="家庭主妇">家庭主妇</option>
<option value="教育/培训">教育/培训</option>
<option value="客户服务/支持">客户服务/支持</option>
<option value="零售商/手工工人">零售商/手工工人</option>
<option value="退休">退休</option>
<option value="无职业">无职业</option>
<option value="销售/市场/广告">销售/市场/广告</option>
<option value="学生">学生</option>
<option value="研究和开发">研究和开发</option>
<option value="一般管理/监督">一般管理/监督</option>
<option value="政府/军队">政府/军队</option>
<option value="执行官/高级管理">执行官/高级管理</option>
<option value="制造/生产/操作">制造/生产/操作</option>
<option value="专业人员">专业人员</option>
<option value="自雇/业主">自雇/业主</option>
<option value="其他">其他</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">　<b>婚姻状况：</b>
<select size=1 name=marital>
<option value="" selected>[请选择]</option>
<option value=未婚>未婚</option>
<option
value=已婚>已婚</option>
<option value=离异>离异</option>
<option
value=丧偶>丧偶</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" class=a4 width="46%">　<b>最高学历：</b>
<select size=1 name=education>
<option value="" selected>[请选择]</option>
<option value=小学>小学</option>
<option
value=初中>初中</option>
<option value=高中>高中</option>
<option
value=大学>大学</option>
<option value=硕士>硕士</option>
<option
value=博士>博士</option>
</select></td>
</tr>
<tr class=a3>
<td height="2" valign="top" width="46%">　<b>毕业院校：</b>
<input type="text" name="college" size="18"></td>
</tr>

<tr class=a3>
<td height="1" valign="top" class=a4 width="46%">　<b>ＱＱ号码： <input type="text" name="userqq" size="18">
</b></td>
</tr>
<tr class=a3>
<td height="1" valign="top" class=a3 width="46%">　<b>ICQ 号码： 
<input type="text" name="icq" size="18">
</b></td>
</tr>
<tr class=a3>
<td valign="top" class=a4 width="46%">　<b>个人主页：</b> 
<input type="text" name="userhome" size="20" value="http://"></td>
</tr>
<tr bgcolor="FFFFFF">
<td colspan="2" height="25" valign="middle" align="left"><b>&nbsp;<img src="images/1.gif">密码保护信息</b>（以下内容建议填写）</td>
</tr>
<tr>
<td class=a4 height="2" align="right" valign="middle" width="46%"><b>密码提示问题：</b></td>
<td class=a4 height="2" align="left" valign="middle" width="54%">
&nbsp;


<select name="question">
<option value="" selected>[请选择]</option>
<option value="最喜欢的宠物？">最喜欢的宠物？</option>
<option value="最喜爱的电影？">最喜爱的电影？</option>
<option value="周年纪念日 [年/月/日]？">周年纪念日 [年/月/日]？</option>
<option value="父亲的名字？">父亲的名字？</option>
<option value="配偶的名字？">配偶的名字？</option>
<option value="第一个孩子的爱称？">第一个孩子的爱称？</option>
<option value="中学的校名？">中学的校名？</option>
<option value="最尊敬的老师？">最尊敬的老师？</option>
<option value="最喜欢的运行队？">最喜欢的运行队？</option>
</select>

</td>
</tr>
<tr class=a3>
<td height="2" align="right" valign="middle" width="46%"><b>密码提示答案：</b></td>
<td height="2" align="left" valign="middle" width="54%"> &nbsp;
<input type="text" name="answer" size="40">
</td>
</tr>

<tr bgcolor="FFFFFF" align="center">
<td height="2" valign="middle" colspan="2">
<input type="submit" name="Submit" value=" 注 册 >>下 一 步 ">
</td>
</tr>
</table>
</td>
</tr></table>

<SCRIPT>valignbottom()</SCRIPT>





<SCRIPT>
language=navigator.language?navigator.language:navigator.browserLanguage
if (language=="zh-cn"){country="中国"}
else
if (language=="zh-hk"){country="中国香港"}
else
if (language=="zh-tw"){country="中国台湾"}
else
if (language=="zh-sg"){country="新加坡"}
else
if (language=="en-us"){country="美国"}
else
if (language=="en-gb"){country="英国"}
else
if (language=="en-au"){country="澳大利亚"}
else
if (language=="en-ca"){country="加拿大"}
else
if (language=="en-nz"){country="新西兰"}
else
if (language=="en-ie"){country="爱尔兰"}
else
if (language=="en-za"){country="南非"}
else
if (language=="en-jm"){country="牙买加"}
else
if (language=="en-bz"){country="伯利兹"}
else
country=""
document.form.country.value=country


function VerifyInput()
{
var newDateObj = new Date()
username=document.form.username.value

//if (/[^\x00-\xff]/g.test(username)){alert("用户名不能含有汉字");return false;}

if (username == "")
{
alert("请输入您的用户名");
document.form.username.focus();
return false;
}

var mail = document.form.usermail.value;
if(mail.indexOf('@',0) == -1 || mail.indexOf('.',0) == -1){
alert("你输入的Email有错误\n请重新检查你的Email");
document.form.usermail.focus();
return false;
}

if (document.form.usermail.value == "")
{
alert("请输入您的EMAIL地址");
document.form.usermail.focus();
return false;
}

return true;
}
</SCRIPT>











<%
htmlend
%>