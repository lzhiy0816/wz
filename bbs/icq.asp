<script>function goodbye(){close();}</script>
<body OnUnload="goodbye()">

<HTML><META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href=images/skins/<%=Request.Cookies("skins")%>/bbs.css rel=stylesheet>
<body topmargin=0 bgcolor="ECE9D8">
<style>
.bt {BORDER-RIGHT: 1px; BORDER-TOP: 1px; FONT-SIZE: 9pt; BORDER-LEFT: 1px; BORDER-BOTTOM: 1px;}
</style><TITLE>����ѶϢ</TITLE>
<SCRIPT>
function check(theForm) {
var mail = document.pagerform.fromemail.value;

if(theForm.from.value == "" ) {
alert("������������֣�");
return false;
}

if(mail.indexOf('@',0) == -1 || mail.indexOf('.',0) == -1){
alert("�������Email�д���\n�����¼�����Email");
document.pagerform.fromemail.focus();
return false;
}

if(theForm.body.value == "" ) {
alert("���ܷ���ѶϢ��");
return false;
}


}
</SCRIPT>
<TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0><TR><form name="pagerform" action="http://wwp.mirabilis.com/scripts/WWPMsg.dll" method="post">
<input type="hidden" name="subject" value="From BBSxp">
<input type="hidden" name="to" value="<%=Request("icq")%>">
<TD height="32">
&nbsp;������֣�<input type="text" size="6" name="from" value="<%=Request.Cookies("username")%>"> ���Email��<input type="text" size="12" name="fromemail">
</TD></TR><TR><TD VALIGN=top ALIGN=right bgcolor="F8F8F8">
<textarea name="body" cols="39" rows="6"></textarea>
</TD></TR></TABLE><TABLE WIDTH=300 BORDER=0 CELLSPACING=0 CELLPADDING=0 height="30">
<tr ALIGN=center><TD><img border=0 src=http://web.icq.com/whitepages/online?img=5&icq=<%=Request("icq")%>>ICQ:<%=Request("icq")%> </td><TD><input OnClick="window.close();" type="reset" name=reset value="ȡ������" ></td>
<TD><input type="submit" value="����ѶϢ" onclick="return check(this.form)"></td>
</TR></form>
</TABLE>