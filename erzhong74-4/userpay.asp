<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chan_const.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dvbbs.stats="������̳��ȯ"
Dvbbs.LoadTemplates("")
Dvbbs.nav()

Dvbbs.Head_var 0,0,"�û��������","usermanager.asp"

If Request("raction")="alipay_return" Then
	AliPay_Return()
	Dvbbs.Footer()
	Response.End
ElseIf Request("action")="alipay_return" Then
	AliPay_Return()
	Dvbbs.Footer()
	Response.End
End If

If Dvbbs.userid=0 Then Dvbbs.AddErrCode(6):Dvbbs.Showerr()
Dvbbs.TrueCheckUserLogin()
CenterMain()
Dvbbs.Showerr()
Dvbbs.Footer()
'Dvbbs.PageEnd()
Sub CenterMain()
%>
	<table border="0" width="<%=Dvbbs.mainsetting(0)%>" cellpadding=2 cellspacing=0 align=center>
		<tr>
		<td width="180" valign=top>
		<%UserInfo()%>
		</td>
		<td width="*" valign=top>
		<%
		Select Case Request.QueryString("action")
			Case "alipay"
				AliPay()
			Case "alipay_1"
				AliPay_1()
			Case "alipay_return"
				AliPay_Return()
			Case "UserCenter"
				UserCenter()
			Case "UserToolsLog_List"
				UserToolsLog_List()
			Case "PayList"
				PayList()
			Case Else
				SmsPayMain()
		End Select
		%>
		</td>
		</tr>
	</table>
<%
End Sub

Sub SmsPayMain()
	MainReadMe(0)

	If Dvbbs.Forum_ChanSetting(3)="0" Then
%>
	<tr><td height=23 class="tablebody2"><B>��������֧�������ȯ</B>��ʹ��ǰ�뵽 <a href="https://www.alipay.com/" target=_blank><font color=red>����Ͱ�.֧����</font></a> ����һ��֧�����˺ţ�֧�����̲���ȡ������</td>
	</tr>
	<FORM TARGET="_blank" METHOD=POST ACTION="?action=alipay">
	<tr><td height=23 class="tablebody1">
	������Ҫ֧���Ľ�
	<input type=text size=5 name="paymoney" value="1" onkeyup="ShowChange(this.value,this,'PAY_M',1)">
	��ȡ<FONT ID="PAY_M" CLASS="REDFONT"><%=CCur(Dvbbs.Forum_ChanSetting(14))*1%></FONT>����̳��ȯ��
	����� 1 Ԫ����� ��
	<input type=submit name=submit value="����֧��">
	</td>
	</tr>
	</FORM>
	<tr><td height=24 class="tablebody1">
	<B>���ɹ�֧������ϵͳ������Ҫ�����ӵ�ʱ��ȴ�֧���������˿����޷�˲�����ˣ�֧���ɹ�����ˢ�´�ҳ�沢�鿴��ȯ���Ƿ���ȷ��</B>
	</td>
	</tr>
	<tr><td height=24 class="tablebody1">
	<iframe src="<%=Dvbbs_Server_Url%>dvbbs/DvDefaultTextAd_1.asp" height=23 width="100%" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=no></iframe>
	</td>
	</tr>
	<%End If%>
	<tr><td height=23 class="tablebody2" style="line-height: 18px"><B>��ȯʹ��С��ʿ</B>��<BR>
	�� ��̳��ȯ�����ڹ�����̳�г��۵ĸ���Ȥζ�Ե���<BR>
	�� ��̳��ȯ�ͽ�ҿ����ڲ�����̳��һЩ��Ҫ��ȯ�����������������������������ȷ�ش������ظ��û��Ȳ���<BR>
	�� ������̳�������䲻ͬ�Ĺ��ܣ��������������Ŀ���û���Ҳ�������Լ����������һЩ��������������Ǯ��ʧ���ߵȣ�<BR>
	�� ��̳��ȯ������̳�û����໥ת�ã�ǰ����Ŀ���û����������̳�����Լ������˵���ת����<BR>
	�� ϵͳ�в�������ĵ��߳�������ʹ�õ�Ŀ�ģ�����Ҫ�û�ͬʱӵ�н�Һ͵�ȯ���ܹ���ģ��в��ֵ���ֻ�������������²Ż���֣��ⲿ�ֵ������õ�ȯ���Ҷ����ܹ��򵽵ġ�</td>
	</tr>
</table>

<SCRIPT LANGUAGE="JavaScript">
<!--
var ProductMoney = <%=Dvbbs.Forum_ChanSetting(14)%>;
function getinfo(v){
	v=parseFloat(v);
	var pag=document.getElementById('pay');
	pag.innerHTML=ProductMoney*v;
}
function ShowChange(Ivalue,Iname,ShowID,Min){
	if(isNaN(Ivalue)){
		Iname.value = Min;
		alert('����д��ȷ����ֵ��');
	}
	else{
		Ivalue = parseFloat(Ivalue);
		Min = parseFloat(Min);
		if (Ivalue<Min){
			Iname.value = Min;
			document.getElementById(ShowID).innerHTML = Min;
			alert('��д��ֵ�������ƣ�');
		}
		else{
			document.getElementById(ShowID).innerHTML = (Ivalue * ProductMoney).toFixed(1);
		}
	}
}
//-->
</SCRIPT>
<%
End Sub

Sub AliPay()
	Dim PayMoney
	PayMoney = Request("paymoney")
	If PayMoney = "" Or Not IsNumeric(PayMoney) Then
		Response.redirect "showerr.asp?ErrCodes=<li>���󣬷Ƿ��ĸ��������&action=OtherErr"
		Exit Sub
	End If
	If PayMoney < 1 Then
		Response.redirect "showerr.asp?ErrCodes=<li>����ÿ�ʶ��������СΪ <B>1</B> Ԫ����ҡ�&action=iOtherErr"
		Exit Sub
	End If
	PayMoney = FormatNumber(PayMoney,2,True,False,False)

	'���ɶ�����:01+yyyyMMddhhmmss+��λ�����
	'���������ִ�
	Dim NowTimes,PayMonth,PayDay,PayHour,PayMin,PaySe,PayDayStr,RandomizeStr,num1
	Dim PayCode,PayCodeEnCode
	NowTimes = Now()
	PayMonth = Month(NowTimes)
	If Len(PayMonth)=1 Then PayMonth = "0" & PayMonth
	PayDay = Day(NowTimes)
	If Len(PayDay)=1 Then PayDay = "0" & PayDay
	PayHour = Hour(NowTimes)
	If Len(PayHour)=1 Then PayHour = "0" & PayHour
	PayMin = Minute(NowTimes)
	If Len(PayMin)=1 Then PayMin = "0" & PayMin
	PaySe = Second(NowTimes)
	If Len(PaySe)=1 Then PaySe = "0" & PaySe
	PayDayStr = Year(NowTimes) & PayMonth & PayDay & PayHour & PayMin & PaySe
	'��������ִ�
	Randomize
	Do While Len(RandomizeStr)<5
		num1 = CStr(Chr((57-48)*rnd+48))
		RandomizeStr = RandomizeStr & num1
	Loop
	'Response.Write RandomizeStr
	'Response.Write "<BR>"
	'Response.Write PayDayStr
	If Dvbbs.Forum_ChanSetting(5) <> "0" Then
		PayCode = "01" & Dvbbs.Forum_ChanSetting(5) & PayDayStr & RandomizeStr
	Else
		PayCode = PayDayStr & RandomizeStr & Left(MD5(Dvbbs.Forum_ChanSetting(4)&Dvbbs.Forum_ChanSetting(6),32),8)
	End If
	Dim EnCodeStr
	
	EnCodeStr="body=Forum points Certificates&notify_url="&Dvbbs_PayTo_Url&"newpay.asp?action=newpay&out_trade_no="&PayCode&"&partner=2088002048522272&payment_type=1&return_url="&Dvbbs_PayTo_Url&"newpay.asp?action=newpay&seller_email="&Lcase(Dvbbs.Forum_ChanSetting(4))&"&service=create_direct_pay_by_user&show_url="&Dvbbs.Get_ScriptNameUrl&"&subject=Forum points Certificates&total_fee="&PayMoney&Dvbbs.Forum_ChanSetting(6)
	EnCodeStr = MD5(EnCodeStr,32)

	'������̳������
	Dvbbs.Execute("InSert Into Dv_ChanOrders (O_type,O_Username,O_isApply,O_issuc,O_PayMoney,O_Paycode,O_AddTime) Values (1,'"&Dvbbs.MemberName&"',0,0,"&PayMoney&",'"&PayCode&"','"&NowTimes&"')")

	'�ύ�������ٷ���������
	If Dvbbs.Forum_ChanSetting(5) <> "0" Then
%>
�����ύ���ݣ����������̳��ַ������URLת������������ȷ������Ϣ�����Ժ󡭡�
<form name="redir" action="<%=Dvbbs_Server_Url%>alipay_t1.aspx?action=pay" method="post">
<INPUT type=hidden name="username" value="<%=Dvbbs.MemberName%>">
<INPUT type=hidden name="paycode" value="<%=PayCode%>">
<INPUT type=hidden name="returnurl" value="<%=Dvbbs.Get_ScriptNameUrl%>UserPay.asp?action=alipay_return">
<INPUT type=hidden name="paymoney" value="<%=PayMoney%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
	Else
%>
�����ύ���ݣ����������̳��ַ������URLת������������ȷ������Ϣ�����Ժ󡭡�
<form name="redir" action="<%=Dvbbs_PayTo_Url%>newpay.asp?action=pay" method="post">
<INPUT type=hidden name="buyer" value="<%=Dvbbs.MemberName%>">
<INPUT type=hidden name="returnurl" value="<%=Dvbbs.Get_ScriptNameUrl%>">
<INPUT type=hidden name="out_trade_no" value="<%=PayCode%>">
<INPUT type=hidden name="seller_email" value="<%=Lcase(Dvbbs.Forum_ChanSetting(4))%>">
<INPUT type=hidden name="total_fee" value="<%=PayMoney%>">
<INPUT type=hidden name="sign" value="<%=EnCodeStr%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
	End If
End Sub

'����֧�����ؽ������������½Ҳ��ִ��
Sub AliPay_Return()
	If Dvbbs.Forum_ChanSetting(5) <> "0" Then
		AliPay_Return_Old()
		Exit sub
	Else
		Dim Rs,Order_No,EnCodeStr,UserInMoney
		Order_No=Dvbbs.Checkstr(Request("out_trade_no"))
		Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_IsSuc=3 And O_PayCode='"&Order_No&"'")
		If not(Rs.Eof And Rs.Bof) Then
			AliPay_Return_Old()
			Exit sub
		End If
		Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_IsSuc=1 And O_PayCode='"&Order_No&"'")
		If not(Rs.Eof And Rs.Bof) Then
			AliPay_Return_Old()
			Exit sub
		End if
		Response.Clear
		Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_IsSuc=0 And O_PayCode='"&Order_No&"'")
		If Rs.Eof And Rs.Bof Then
			Response.Write "fail"
		Else
			Response.Write "success"
			Dvbbs.Execute("Update Dv_ChanOrders Set O_IsSuc=3 Where O_ID = " & Rs("O_ID"))
		End If
		Response.End
	End If
End Sub

Sub AliPay_Return_Old()		
	'�õ����жϷ��ز���
	Dim PayCode,SignStr,Success,UserInMoney
	PayCode = Replace(Request("out_trade_no"),"'","")
	Success = Request("is_success")
	If PayCode = "" Or Success = "" Then
		Response.redirect "showerr.asp?ErrCodes=<li>���󣬷Ƿ��Ķ���������&action=OtherErr"
		Exit Sub
	End If
	If Success<>"T" Then
		Response.redirect "showerr.asp?ErrCodes=<li>����֧��ʧ�ܣ�����ϸ�������֧����Ϣ��<a href=""UserPay.asp"">���½���֧��ҳ��</a>��&action=iOtherErr"
		Exit Sub
	End If

	'��֤������Ϣ
	Dim Rs
	Set Rs = Dvbbs.Execute("Select * From [Dv_ChanOrders] Where O_PayCode='"&PayCode&"'")
	If Rs.Eof And Rs.Bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>�����Ҳ����ö�����Ϣ��ö�����֧���ɹ���&action=OtherErr"
		Exit Sub
	Else
		If CInt(rs("O_issuc"))=3 Or CInt(rs("O_issuc"))=1 Then
			dim alipayNotifyURL,ResponseTxt,Retrieval
			alipayNotifyURL="http://notify.alipay.com/trade/notify_query.do?partner=2088002048522272&notify_id="&request("notify_id")
			Set Retrieval=Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
			Retrieval.setOption 2,13056 
			Retrieval.open "GET",alipayNotifyURL,False,"","" 
			Retrieval.send()
			ResponseTxt=Retrieval.ResponseText
			Set Retrieval=Nothing
			If ResponseTxt="false" Then Response.redirect "showerr.asp?ErrCodes=<li>���󣬷Ƿ��Ķ���������&action=OtherErr":Exit Sub
			'�������ݿ�����
			UserInMoney = Rs("O_PayMoney")
			If CInt(rs("O_issuc"))=3 then
			'�����û�����
			Dvbbs.Execute("Update Dv_User Set UserTicket = UserTicket + " & Dvbbs.Forum_ChanSetting(14) * UserInMoney & " Where UserName='"&Rs("O_UserName")&"'")
			If Dvbbs.UserID > 0 And Lcase(Dvbbs.MemberName)=Lcase(Rs("O_UserName")) Then
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) + cCur(Dvbbs.Forum_ChanSetting(14) * UserInMoney)
			End If
			'���¶���״̬
			Dvbbs.Execute("Update Dv_ChanOrders Set O_IsSuc=1 Where O_ID = " & Rs("O_ID"))
			End if
		Else
			Response.redirect "showerr.asp?ErrCodes=<li>�����Ҳ����ö�����Ϣ��ö�����֧���ɹ���&action=OtherErr"
			Exit Sub
		End if
	End If
	Rs.Close
	Set Rs=Nothing
%>
<!--��̳�����ɹ���Ϣ-->
<br>
<table cellpadding=0 cellspacing=1 align=center class="tableborder1" style="width:75%">
<tr align=center>
<th width="100%">��̳�ɹ���Ϣ
</td>
</tr>
<tr>
<td width="100%" class="tablebody1">
<b>�����ɹ���</b><br><br>
<li>�ɹ��������ζһ��� <B><font color=red><%=(Dvbbs.Forum_ChanSetting(14) * UserInMoney)%></font></B> ����̳��ȯ��
</td></tr>
<tr align=center><td width="100%" class="tablebody2">
<a href="usermanager.asp"> << �����û��������</a> &nbsp;&nbsp;||&nbsp;&nbsp; <a href="UserPay.asp?action=UserCenter"> ȥ�ѵ�ȯת������̳���>></a> 
</td></tr>
</table><br>
<%
End sub
'--------------------------------------------------------------------------------
'�û���Ϣ
'--------------------------------------------------------------------------------
Sub UserInfo()
	Dim Sql,Rs,UserToolsCount
	'Sql = "Select Sum(ToolsCount) From [Dv_Plus_Tools_Buss] where UserID="& Dvbbs.UserID
	'Set Rs = Dvbbs.Plus_Execute(Sql)
	'UserToolsCount = Rs(0)
	'If IsNull(UserToolsCount) Then UserToolsCount = 0
%>
<table cellpadding="0" cellspacing="1" align="center" class="tableborder1" Style="Width:100%">
	<tr>
		<th height=23 >��������</th>
	</tr>
	<tr>
		<td align=center class="tablebody1">
			<table border="0" cellpadding="0" cellspacing="1" align="center" Style="Width:90%">
				<tr>
					<td class="tablebody2" style="text-align:left;">��ң�
						<B>
							<font color="<%=Dvbbs.mainsetting(1)%>">
								<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text %>
							</font>
						</B> ��
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left;">��ȯ��<B>
						<font color="<%=Dvbbs.mainsetting(1)%>">
							<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%>
						</font></B> ��
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left;">��Ǯ��<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%></td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left;">���£�<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text%></td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left;">���飺<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%></td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left;">������<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%></td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left;">������<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text%></td>
				</tr>
				<tr><td class="tablebody1"></td></tr>
			</table>
		</td>
	</tr>
</table>
<%
End Sub

'--------------------------------------------------------------------------------
'���ת��
'--------------------------------------------------------------------------------
Sub UserCenter()
	If Request("react") = "Savechange" Then
		If Not Dvbbs.ChkPost() Then Dvbbs.AddErrCode(16):Dvbbs.Showerr()
		Dim userWealth,userep,usercp,userticket,UpUserMoney
		Dim Sql,Rs
		userWealth = Dvbbs.CheckNumeric(Request.Form("userWealth"))
		userep = Dvbbs.CheckNumeric(Request.Form("userep"))
		usercp = Dvbbs.CheckNumeric(Request.Form("usercp"))
		userticket = Dvbbs.CheckNumeric(Request.Form("userticket"))
		UpUserMoney = 0
		If userWealth<0 or userep<0 or usercp<0 or userticket<0 Then Dvbbs.AddErrCode(35):Dvbbs.Showerr()

		Dim ErrMsg
		If userWealth>0 And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)<CCur(Dvbbs.Forum_setting(93)) Then ErrMsg="��Ľ�Ǯ���㡣"
		If userep>0 And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)<CCur(Dvbbs.Forum_setting(94)) Then ErrMsg="��ľ��鲻�㡣"
		If usercp>0 And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)<CCur(Dvbbs.Forum_setting(95)) Then ErrMsg="����������㡣"
		If userticket And CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)<CCur(Dvbbs.Forum_setting(96)) Then ErrMsg="��ĵ�ȯ���㡣"
		If Trim(ErrMsg)<>"" Then
			Response.redirect "showerr.asp?ErrCodes=<li>"&ErrMsg&"&action=OtherErr"
		End If

		If userWealth>=1 and userWealth<=CCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text) and cCur(Dvbbs.Forum_setting(93))<>0 Then
			If Cint(userWealth / cCur(Dvbbs.Forum_setting(93))) > 0 Then
				UpUserMoney = UpUserMoney + Cint(userWealth / cCur(Dvbbs.Forum_setting(93)))
				userWealth = Cint(userWealth / cCur(Dvbbs.Forum_setting(93))) * cCur(Dvbbs.Forum_setting(93))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text) - userWealth
			Else
				userWealth = 0
			End If
		Else
			userWealth = 0
		End If

		If userep>=1 and userep<=cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text) and cCur(Dvbbs.Forum_setting(94))<>0 Then
			If Cint(userep / cCur(Dvbbs.Forum_setting(94))) > 0 Then
				UpUserMoney = UpUserMoney + Cint(userep / cCur(Dvbbs.Forum_setting(94)))
				userep = Cint(userep / cCur(Dvbbs.Forum_setting(94))) * cCur(Dvbbs.Forum_setting(94))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text) - userep
			Else
				userep = 0
			End If
		Else
			userep = 0
		End If
		If usercp>=1 and usercp<=cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text) and cCur(Dvbbs.Forum_setting(95))<>0 Then
			If Cint(usercp / cCur(Dvbbs.Forum_setting(95))) > 0 Then
				UpUserMoney = UpUserMoney + Cint(usercp / cCur(Dvbbs.Forum_setting(95)))
				usercp = Cint(usercp / cCur(Dvbbs.Forum_setting(95))) * cCur(Dvbbs.Forum_setting(95))
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text) - usercp
			Else
				usercp = 0
			End If
		Else
			usercp = 0
		End If
		If userticket>=1 and userticket<=cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) and Dvbbs.Forum_setting(96) <> 0 Then
			Userticket = Clng(Userticket)
			If Cint(userticket / Dvbbs.Forum_setting(96)) > 0 Then
				UpUserMoney = UpUserMoney + Cint(userticket / Dvbbs.Forum_setting(96))
				userticket = Cint(userticket / Dvbbs.Forum_setting(96)) * Dvbbs.Forum_setting(96)
				Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text) - userticket
			Else
				userticket = 0
			End If
		Else
			userticket = 0
		End If
		If UpUserMoney < 1 Then 
			 Response.redirect "showerr.asp?ErrCodes=<li>����дת�������ݻ��õĽ����̫�٣�&action=OtherErr"
		Else
			Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text  = cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text ) + UpUserMoney
			Sql = "Update Dv_user set userWealth = "&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text&",userEP="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text&",userCP="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text&",UserMoney="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &",UserTicket="&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text&" where UserID="&Dvbbs.UserID
			Dvbbs.Execute(Sql)
			Dim LogMsg
			LogMsg = "���ת���ɹ�������ܽ����Ϊ<b>"&UpUserMoney&"</b>,��Ǯ����<b>"&userWealth&"</b>,�������<b>"&userep&"</b>,��������<b>"&usercp&"</b>,��ȯ����<b>"&userticket&"</b>��"
			'Call Dvbbs.ToolsLog(0,0,0,0,0,LogMsg,Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text &"|"&Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)
			Dvbbs.Dvbbs_Suc(LogMsg)
		End If
	Else
%>
	<table border=0 cellpadding=3 cellspacing=1 class="tableborder1" align=center style="width:100%">
	<tr><th height=20 colspan="5">��̳���ת��</th></tr>
	<tr><td height=20 colspan="5" class="tablebody1"><li>�����û�����Ǯ�����顢��������ȯת���ɽ�ҡ�</td></tr>
    <tr>
      <th width="30%" height="20">���ת������</th>
      <th width="15%">ת����Ŀ</th>
	  <th width="20%">ת����Ϣ</th>
      <th width="15%">ת������</th>
	  <th width="20%">ת�����ý��</th>
    </tr>
	<form action="UserPay.asp?action=UserCenter&react=Savechange" method=post NAME=CenterForm>
    <tr>
      <td rowspan="5" class="tablebody1">
		<table border="0" cellpadding=0 cellspacing=1 align=center Style="Width:90%">
			<tr><td class="tablebody1">&nbsp;&nbsp;&nbsp;&nbsp;<a href="UserPay.asp"><font color=red>ǰ��������̳��ȯ</font></a></td></tr>
			<tr><td class="tablebody2">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> ��� = <font class=redfont><%=Dvbbs.Forum_setting(93)%></font> ��Ǯ</b></td></tr>
			<tr><td class="tablebody1">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> ��� = <font class=redfont><%=Dvbbs.Forum_setting(94)%></font> ����</b></td></tr>
			<tr><td class="tablebody2">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> ��� = <font class=redfont><%=Dvbbs.Forum_setting(95)%></font> ����</b></td></tr>
			<tr><td class="tablebody1">&nbsp;&nbsp;&nbsp;&nbsp;<b><font class=redfont>1</font> ��� = <font class=redfont><%=Dvbbs.Forum_setting(96)%></font> ��ȯ</b></td></tr>
			<tr><td class="tablebody2"></td></tr>
		</table>
	  </td>
      <td class="tablebody2" align=center>ӵ�н�Ǯֵ��</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="userWealth" value="0" onkeyup="ShowChange(this.value,this,'Show_Money',<%=Dvbbs.Forum_setting(93)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%>)"></td>
	  <td class="tablebody1" ID=Show_Money>0</td>
    </tr>
    <tr>
      <td class="tablebody2" align=center>ӵ�о���ֵ��</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="userep" value="0" onkeyup="ShowChange(this.value,this,'Show_EP',<%=Dvbbs.Forum_setting(94)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%>)"></td>
	  <td class="tablebody1" ID=Show_EP>0</td>
    </tr>
    <tr>
      <td class="tablebody2" align=center>ӵ������ֵ��</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="usercp" value="0" onkeyup="ShowChange(this.value,this,'Show_CP',<%=Dvbbs.Forum_setting(95)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%>)"></td>
	  <td class="tablebody1" ID=Show_CP>0</td>
    </tr>
    <tr>
      <td class="tablebody2" align=center>ӵ�е�ȯֵ��</td>
      <td class="tablebody1"><font class=redfont><%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%></font></td>
	  <td class="tablebody1"><INPUT TYPE="text" NAME="userticket" value="0" onkeyup="ShowChange(this.value,this,'Show_Ticket',<%=Dvbbs.Forum_setting(96)%>,<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%>)"></td>
	  <td class="tablebody1" ID=Show_Ticket>0</td>
    </tr>
	<tr>
      <td class="tablebody2" align=center colspan="4">
	  <INPUT TYPE="submit" value="ȷ��ת��">&nbsp;&nbsp;<INPUT TYPE="reset" value="��������"></td>
    </tr>
	</form>
	</table>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function ShowChange(Ivalue,Iname,ShowID,Sys,User){
		if(isNaN(Ivalue)){
			Iname.value = 0;
			alert('����д��ȷ����ֵ��');
		}
		else{
			Ivalue = parseFloat(Ivalue);
			Sys = parseFloat(Sys);
			User = parseFloat(User);
			if (Ivalue>User||Ivalue<0){
				Iname.value = 0;
				document.getElementById(ShowID).innerHTML = 0;
				alert('��д��ֵ�������ƣ�');
			}
			else{
				document.getElementById(ShowID).innerHTML = (Ivalue / Sys).toFixed(1);
			}
		}
	}
	//-->
	</SCRIPT>
<%
	End If
End Sub

'�û������б�
Sub PayList()
	Dim Success
	Success = Dvbbs.CheckNumeric(Request("Suc"))

	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString
	PageSearch = "action=PayList&Suc=" & Success
	Endpage = 0
	MaxRows = 20
	Page = Request("Page")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Response.Write "<script language=""JavaScript"" src=""inc/Pagination.js""></script>"

	MainReadMe(1)
%>
		</td>
		</tr>
		<tr><td colspan=3><hr style="BORDER: #807d76 1px dotted;height:1px;">


<table border="0" cellpadding=3 cellspacing=1 align=center class="tableborder1" style="width:100%">
	<tr><td height=23 class="tablebody2" colspan=6 style="line-height: 18px">
<%
	Dim Rs,Sql
	Select Case Success
	Case 0
		Response.Write Dvbbs.MemberName & " ��������̳����֧�����׶���"
		Sql = "Select O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID From Dv_ChanOrders Where O_UserName = '"&Dvbbs.MemberName&"' Order By O_AddTime Desc"
	Case 1
		Response.Write Dvbbs.MemberName & " ��������̳����֧�����׳ɹ�����"
		Sql = "Select O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID From Dv_ChanOrders Where O_IsSuc = 1 And O_UserName = '"&Dvbbs.MemberName&"' Order By O_AddTime Desc"
	Case 2
		Response.Write Dvbbs.MemberName & " ��������̳����֧������ʧ�ܶ���"
		Sql = "Select O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID From Dv_ChanOrders Where O_IsSuc = 0 And O_UserName = '"&Dvbbs.MemberName&"' Order By O_AddTime Desc"
	End Select
%>
	</td></tr>
	<tr>
	<th height=23 width="15%">��������</th>
	<th width="20%">������</th>
	<th width="15%">֧�����</th>
	<th width="15%">����״̬</th>
	<th width="15%">����ʱ��</th>
	<th width="20%">����</th>
	</tr>
<%
	Dim i
	Set Rs = server.CreateObject ("adodb.recordset")
	If Not IsObject(Conn) Then ConnectionDatabase
	Rs.Open Sql,Conn,1,1
	If Rs.Eof And Rs.Bof Then
		Response.Write "<tr><td height=23 class=""tablebody1"" colspan=6>��ǰ��û�ж�����</td></tr>"
		Response.Write "</table>"
	Else
		CountNum = Rs.RecordCount
		If CountNum Mod MaxRows=0 Then
			Endpage = CountNum \ MaxRows
		Else
			Endpage = CountNum \ MaxRows+1
		End If
		Rs.MoveFirst
		If Page > Endpage Then Page = Endpage
		If Page < 1 Then Page = 1
		If Page >1 Then 				
			Rs.Move (Page-1) * MaxRows
		End if
		SQL=Rs.GetRows(MaxRows)
		'O_Type,O_PayCode,O_PayMoney,O_IsSuc,O_AddTime,O_ID
		For i=0 To Ubound(SQL,2)
%>
	<tr align=center>
	<td height=23 class="tablebody1">
<%
	Select Case SQL(0,i)
	Case 1
		Response.Write "����֧��"
	Case Else
		Response.Write "<font color=gray>δ֪</font>"
	End Select
%>
	</td>
	<td class="tablebody1"><%=SQL(1,i)%></td>
	<td class="tablebody1"><%=SQL(2,i)%></td>
	<td class="tablebody1">
<%
	Select Case SQL(3,i)
	Case 0
		Response.Write "<font color=gray>ʧ��</font>"
	Case 1
		Response.Write "�ɹ�"
	Case Else
		Response.Write "<font color=gray>δ֪</font>"
	End Select
%>
	</td>
	<td class="tablebody1"><%=SQL(4,i)%></td>
	<td class="tablebody1">&nbsp;
	</td>
	</tr>
<%
		Next
	Response.Write "</table>"
	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
	End If
	Rs.Close
	Set Rs=Nothing

End Sub

'���»�ý���״̬
Sub AliPay_1()
	Dim ID,Rs
	Dim PayMoney,PayCode
	ID = Request("ID")
	If ID = "" Or Not IsNumeric(ID) Then
		Response.redirect "showerr.asp?ErrCodes=<li>���󣬷Ƿ��Ķ���������&action=OtherErr"
		Exit Sub
	Else
		ID = cCur(ID)
	End If
	Set Rs = Dvbbs.Execute("Select * From Dv_ChanOrders Where O_ID = "&ID&" And O_UserName = '"&Dvbbs.MemberName&"'")
	If Rs.Eof And Rs.Bof Then
		Response.redirect "showerr.asp?ErrCodes=<li>�����Ҳ�����صĶ�����Ϣ��&action=OtherErr"
		Exit Sub
	Else
		PayMoney = Rs("O_PayMoney")
		PayMoney = FormatNumber(PayMoney,2,True,False,False)
		PayCode = Rs("O_PayCode")
	End If
	Rs.Close
	Set Rs=Nothing
	'�ύ�������ٷ���������
%>
�����ύ���ݣ����������̳��ַ������URLת������������ȷ������Ϣ�����Ժ󡭡�
<form name="redir" action="<%=Dvbbs_Server_Url%>alipay_t1.aspx?action=pay_1" method="post">
<INPUT type=hidden name="username" value="<%=Dvbbs.MemberName%>">
<INPUT type=hidden name="paycode" value="<%=PayCode%>">
<INPUT type=hidden name="returnurl" value="<%=Dvbbs.Get_ScriptNameUrl%>UserPay.asp?action=alipay_return">
<INPUT type=hidden name="paymoney" value="<%=PayMoney%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
<%
End Sub

Sub UserToolsLog_List()

	Dim Rs,Sql,i,LogType
	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString
	LogType = "δ֪|ʹ��|ת��|��ֵ|����|����|VIP����"
	LogType = Split(LogType,"|")
	PageSearch = "action=UserToolsLog_List"
	Endpage = 0
	MaxRows = 20
	Page = Request("Page")
	If IsNumeric(Page) = 0 or Page="" Then Page=1
	Page = Clng(Page)
	Response.Write "<script language=""JavaScript"" src=""inc/Pagination.js""></script>"

	If Request.QueryString("UserID")<>"" and IsNumeric(Request.QueryString("UserID")) Then _
	SqlString = "and UserID="&Dvbbs.CheckNumeric(Request.QueryString("UserID"))

	MainReadMe(1)
%>
		</td>
		</tr>
		<tr><td colspan=3><hr style="BORDER: #807d76 1px dotted;height:1px;">
<table border="0" cellpadding=3 cellspacing=1 align=center class="tableborder1" Style="Width:100%">
	<tr>
	<th height=23 width="15%">��������</th>
	<th width="10%">����</th>
	<th width="*%">��������</th>
	<th width="5%">���</th>
	<th width="5%">��ȯ</th>
	<th width="5%">����</th>
	<th width="13%">ʹ��IP</th>
	<th width="12%">ʱ��</th>
	</tr>
<%
	Dim ToolsNames
	Dvbbs.forum_setting(90)=0
	If Dvbbs.forum_setting(90)="1" Then
		Set Rs = Dvbbs.Plus_Execute("Select ID,ToolsName From Dv_Plus_Tools_Info Order By ID")
		If Not (Rs.Eof And Rs.Bof) Then
			Sql = Rs.GetRows(-1)
		End If
		Rs.Close
		Set ToolsNames = server.CreateObject ("adodb.recordset")
		For i=0 to Ubound(Sql,2)
			ToolsNames.add Sql(0,i),Sql(1,i)
		Next
		ToolsNames.add -88,"ħ�������ͷ��"		'���ӵ�����ħ�������ͷ��IDΪ-88
	End If

	'T.ToolsName=0,L.CountNum=1,L.Log_Money=2,L.Log_Ticket=3,L.Log_IP=4,L.Log_Time=5,L.Log_Type=6,L.Conect=7
	Sql = "Select ToolsID,CountNum,Log_Money,Log_Ticket,Log_IP,Log_Time,Log_Type,Conect From Dv_MoneyLog Where AddUserID="&Dvbbs.UserID&" And Not BoardID=-1 Order By Log_Time Desc"
	'Response.Write Sql
	Set Rs = server.CreateObject ("adodb.recordset")
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,1
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,conn,1,1
	End If

	If Not (Rs.Eof And Rs.Bof) Then
		CountNum = Rs.RecordCount
		If CountNum Mod MaxRows=0 Then
			Endpage = CountNum \ MaxRows
		Else
			Endpage = CountNum \ MaxRows+1
		End If
		Rs.MoveFirst
		If Page > Endpage Then Page = Endpage
		If Page < 1 Then Page = 1
		If Page >1 Then 				
			Rs.Move (Page-1) * MaxRows
		End if
		SQL=Rs.GetRows(MaxRows)
	Else
		Response.Write "<tr><td class=""Tablebody1"" colspan=""8"" align=center>���߻�δ���ӣ�</td></tr></table>"
		Exit Sub
	End If
	Rs.close:Set Rs = Nothing
	
	'��������б�
	For i=0 To Ubound(SQL,2)
%>
	<tr>
	<td class="Tablebody1" align=center height=24>
<%
	If Dvbbs.forum_setting(90)="1" Then
		Response.Write ToolsNames(SQL(0,i))
	Else
		Response.Write "<font color=gray>δ֪</font>"
	End If
%>
	</td>
	<td class="Tablebody1" align=center><%=LogType(SQL(6,i))%></td>
	<td class="Tablebody1"><%=SQL(7,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(2,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(3,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(1,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(4,i)%></td>
	<td class="Tablebody1" align=center><%=SQL(5,i)%></td>
	</tr>
<%
	Next
	Set ToolsNames = Nothing
	Response.Write "</table>"
	PageSearch=Replace(Replace(PageSearch,"\","\\"),"""","\""")
	Response.Write "<SCRIPT>PageList("&Page&",3,"&MaxRows&","&CountNum&","""&PageSearch&""",1);</SCRIPT>"
End Sub

Sub MainReadMe(str)
%>
<table border="0" cellpadding=0 cellspacing=1 align=center class="tableborder1" Style="Width:100%">
	<tr>
	<th height=23>������̳��ȯ</th></tr>
	<tr><td height=24 class="tablebody2" align=center><a href="?action=PayList">���н��׼�¼</a> | <a href="?action=PayList&Suc=1">�ѳɹ�����</a> | <a href="?action=PayList&Suc=2">δ�ɹ�����</a> | <a href="?action=UserToolsLog_List">��һ��ȯʹ�ü�¼</a> | <a href="?action=UserCenter"><font color=red>�һ���̳���</font></a> | <a href="UserPay.asp"><font color=red>������̳��ȯ</font></a></td>
	</tr>
	<tr><td height=23 class="tablebody1" style="line-height: 18px"><B>˵��</B>��<BR>
	�� ͨ������֧���ɻ�<font color=red>����</font>��Ӧ����̳��ȯ<BR>
	�� ÿͨ������֧�� <font color=red><B>1</B></font> Ԫ�ɻ��� <font color=red><B><%=Dvbbs.Forum_ChanSetting(14)%></B></font> ����̳��ȯ<BR>
	�� ��̳��ȯ�����ã��ɹ�����̳�и���Ȥζ���ߣ����ܸ�����Ȥ����̳����<BR>
	�� ��ȯ�Ļ�ȡ���̣�����������ʾѡ������֧����ͨ������֧���ɹ��Ľ���ֱ�Ӷ�����̳�˺Ž�����Ӧ�ĵ�ȯ<BR>
	</td>
	</tr>
<%
	If Str = 1 Then Response.Write "</table>"
End Sub

Function URLDecode(enStr)
	dim deStr
	dim c,i,v
	deStr=""
	for i=1 to len(enStr)
		c=Mid(enStr,i,1)
		if c="%" then
			v=eval("&h"+Mid(enStr,i+1,2))
			if v<128 then
				deStr=deStr&chr(v)
				i=i+2
			else
				if isvalidhex(mid(enstr,i,3)) then
				if isvalidhex(mid(enstr,i+3,3)) then
					v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
					deStr=deStr&chr(v)
					i=i+5
				else
					v=eval("&h"+Mid(enStr,i+1,2)+cstr(hex(asc(Mid(enStr,i+3,1)))))
					deStr=deStr&chr(v)
					i=i+3 
				end if 
				else 
					destr=destr&c
				end if
			end if
		else
			if c="+" then
				deStr=deStr&" "
			else
				deStr=deStr&c
			end if
		end if
	next
	URLDecode=deStr
End Function

function isvalidhex(str)
	dim c
	isvalidhex=true
	str=ucase(str)
	if len(str)<>3 then isvalidhex=false:exit function
	if left(str,1)<>"%" then isvalidhex=false:exit function
	c=mid(str,2,1)
	if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex=false:exit function
	c=mid(str,3,1)
	if not (((c>="0") and (c<="9")) or ((c>="A") and (c<="Z"))) then isvalidhex=false:exit function
end function

%>