<!-- #include file="setup.asp" -->
<%

top
if Request.Cookies("username")=empty then error("<li>����δ<a href=login.asp>��¼</a>����")

money=int(Request("money"))

if Request("menu")="win" then
if instr(Request.ServerVariables("http_referer"),"http://"&Request.ServerVariables("server_name")&""&Request.ServerVariables("script_name")&"") = 0 then error("<li>��Դ����")
if Request.Cookies("gobang")=empty or Request.Cookies("gobang")>10 or Request.Cookies("sessionid")<>session.sessionid then error("<li>�Ƿ�����")
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
rs("money")=rs("money")+Request.Cookies("gobang")*2
rs.update
rs.close
Response.Cookies("gobang")=""
response.redirect "loading.htm"

elseif Request("menu")="lost" then
if Request.Cookies("gobang")=empty or Request.Cookies("gobang")>10 or Request.Cookies("sessionid")<>session.sessionid then error("<li>�Ƿ�����")
Response.Cookies("gobang")=""
response.redirect "loading.htm"
end if


%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title>BBSxp������</title>

<style>TABLE{BORDER-TOP:0px;BORDER-LEFT:0px;BORDER-BOTTOM:1px}TD{BORDER-RIGHT:0px;BORDER-TOP:0px}
</style>
<table width=97% align="center" border="0">
<tr>
<td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
<td>��<img src="images/closedfold.gif" border="0">��<a href=main.asp><%=clubname%></a><br>
��<img src="images/bar.gif" border="0"><img src="images/openfold.gif" border="0">��<a href="Gobang.asp">�� �� ��</a></td>
</tr>
</table><br>
<center>
<%


if Request("menu")="" then
%><table cellSpacing="1" cellPadding="2" width="80%" border="0" class=a2>
	<tr class=a1>
		<td colspan="5" height="25" align="center"><strong>��Ϸ����</strong></td>
	</tr>
	<tr class=a3>
		<td colspan="5"><p>����������Դ���й��Ŵ��Ĵ�ͳ�ڰ�����֮һ���ִ����������ĳ�֮Ϊ �� ���� �� 
		��Ӣ��Ϊ ��Renju�� ��Ӣ�ĳ�֮Ϊ ��Gobang�� �� ��FIR��(Five in a Row ����д ) ������ �� ������ �� �� �� 
		������ �� �� �� ���� �� �� �� ��Ŀ �� �� �� ��Ŀ�� �� �� �� ��� �� �ȶ��ֳ�ν��</p>
		<p>�����岻������ǿ˼ά������������������Ҹ��������������������ԡ�����������ִ����е��������� �� �̡�ƽ���� �� 
		�����йŵ���ѧ�ĸ���ѧ�� �� �������� �� 
		�������м���ѧ�����ԣ�Ϊ����Ⱥ����ϲ���ּ���������µļ��ɺ͸�ˮƽ�Ĺ����Ա������������Ļ�ԴԨ���������ж��������غ�������ֱ�ۣ����� �� �� 
		�� �ĸ������ �� �� �� �����ӡ����������Ļ��Ľ����㣬�ǹŽ�����Ľᾧ�� </p>
		<p>��������Դ�ڹŴ��й�����չ���ձ���������ŷ�ޡ���������Χ��Ĺ�ϵ������˵����һ˵����Χ�壬���� �� Ң��Χ�� �� 
		֮ǰ������������������Ϸ��һ˵Դ��Χ�壬��Χ�巢չ��һ����֧�����й����Ļ���������ǵ��������Ŵ���������������Χ����ͬ���ݺ��ʮ�ߵ����������Լ��Χ��һ�����ҹ��ϱ���ʱ�Ⱥ��볯�ʡ��ձ��ȵء����ձ�ʷ�����׽��ܣ��й��Ŵ����������Ǿ��ɸ��� 
		( ���� ) ���� 1688 ���� 1704 ����ձ�Ԫ»ʱ�������ձ��ġ����ձ����� 32 �� ( ��Ԫ 1899 �� ) ���������������� �� 
		���� �� ��һ���Ʋű���ʽȷ��������ȡ���� �� ������ϱڣ����������� �� ���Ӵˣ����������˲��ϵĸ�������Ҫ�ǹ���ı仯 ( 
		����ִ����һ�������� ) �����磬 1899 ��涨����ֹ�ڰ�˫���� �� ˫�� �� �� 1903 ��涨��ֻ��ֹ�ڷ��� �� ˫�� �� �� 
		1912 ��涨���ڷ������� �� ˫�� �� �����䣻 1916 ��涨���ڷ������� �� ���� �� �� 1918 ��涨���ڷ������� �� 
		�ġ������� �� �� 1931 ��涨���ڷ������� �� ˫�� �� �����涨�� 19��19 ��Χ���̸�Ϊ 15��15 
		������ר�����̡������ͳ������崫��ŷ�޲�Ѹ�ٷ���ȫŷ��ͨ��һϵ�еı仯��ʹ��������һ�򵥵���Ϸ���ӻ����淶���������ճ�Ϊ�����ְҵ���������壬ͬʱҲ��Ϊһ�ֹ��ʱ����塣 
		</p>
		</td>
	</tr>
	<tr class=a1>
		<td colspan="5" align="center" height="25">
		<p><strong>��Ϸ���� </strong></p>
		</td>
	</tr>
	<tr class=a3>
		<td colspan="5">1 ������������������֮����еľ��������Һ������С�<br>
		2 ��������ר����Ϊ 15��15 �������ӵķ���Ϊ�ᡢ����б��<br>
				3 �����������κ�һ���γ�����Ϊʤ��</td>
	</tr>
	<tr class=a4>
		<td width="30%" valign="bottom" align="center" height="25">
		<p><b>��ѡ����Ҫ�µĶ�ע��</b></td>
		<td width=15% valign=bottom align="center" height="25"><b><a href=?menu=start&money=1>1���</a></b></td>
		<td width="15%" valign="bottom" align="center" height="25"><b><a href="?menu=start&money=2">2���</a></b></td>
		<td width="15%" valign="bottom" align="center" height="25"><b><a href="?menu=start&money=5">5���</a></b></td>
		<td width="15%" valign="bottom" align="center" height="25"><b><a href="?menu=start&money=10">10���</td>
	</tr>
	

	</table>

<%



else
if money > 10 or money < 1 then  error2("���Ľ���������")
sql="select * from [user] where username='"&Request.Cookies("username")&"'"
rs.Open sql,Conn,1,3
if rs("money") < money then error2("���Ľ�Ҳ��㣡")
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

	<tr><td align="center">��</td><td align="center">
		����ǰ���� <%=nowmoney%>  ���</td>
		<td align="center">���ֶ�ע��<%=money%> ���</td><td align="center">��</td></tr>
	<tr>
		<td align="center">
<img border="0" src="images/plus/Gobang1.gif"><br>
<br>
<img src="<%=userface%>"></td>



		<td background="images/plus/Gobangbg.gif" width="500" colspan="2">
		
		<div id="losshtml" style="position:absolute;visibility:hidden;"><table width=500 height=140 border=0 cellspacing=2 cellpadding=0><tr><td align=center><b><font size="7" color="#478130">�����飬</font><font size="7" color="#FFFFFF">��</font><font size="7" color="#478130">ʤ!<br>
			</font><font size="3" color="#478130">���ź���������</font> <font size="3" color="#FF0000"><%=Request("money")%></font> <font size="3" color="#478130">���.</font></b></td></tr></table></div>

		


<div id="winhtml" style="position:absolute;visibility:hidden;"><table width=500 height=140 border=0 cellspacing=2 cellpadding=0><tr><td align=center><b><font size="7" color="#478130">�����飬</font><font size="7">��</font><font size="7" color="#478130">ʤ!<br>
	</font><font color="#478130">��ϲ���������� </font><font size="3" color="#FF0000"><%=Request("money")%></font> <font size="3" color="#478130">���.</font></b></td></tr></table></div>

		
		
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