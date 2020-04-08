<%
Dim Dv_Tools
Set Dv_Tools=new Plus_Tools_Cls

Class Plus_Tools_Cls
	Public ToolsID,ToolsInfo,ToUserInfo,UserToolsInfo,ToolsSetting
	Private buyCount

	Private Sub Class_Initialize()
		buyCount = 1
		ToolsID = CheckNumeric(Request("ToolsID"))
		If DVbbs.Forum_Setting(90)=0 and IsEmpty(session("flag")) Then ShowErr(1)	'�����ѹر�
	End Sub

	Public Sub ChkToolsLogin()
		If Dvbbs.UserID=0 Then Dvbbs.AddErrCode(6):Dvbbs.Showerr()	'�ж��û��Ƿ����ߡ�
		If ToolsID=0 Then ShowErr(3):Exit Sub
		GetToolsInfo	'��ȡ����������Ϣ
	End Sub

	'---------------------------------------------------
	'��ȡ����ϵͳ��Ϣ
	'---------------------------------------------------
	Private Sub GetToolsInfo()
		Dim Sql,Rs
		'ID=0 ,ToolsName=1 ,ToolsInfo=2 ,IsStar=3 ,SysStock=4 ,UserStock=5 ,UserMoney=6 ,UserPost=7 ,UserWealth=8 ,UserEp=9 ,UserCp=10 ,UserGroupID=11 ,boardID=12,UserTicket=13,buyType=14,ToolsImg=15,ToolsSetting=16
		Sql = "Select ID,ToolsName,ToolsInfo,IsStar,SysStock,UserStock,UserMoney,UserPost,UserWealth,UserEp,UserCp,UserGroupID,boardID,UserTicket,buyType,ToolsImg,ToolsSetting From [Dv_Plus_Tools_Info] Where ID="& ToolsID
		Set Rs = Dvbbs.Plus_Execute(Sql)
		If Rs.Eof Then
			ShowErr(3):Exit Sub
		Else
			Sql = Rs.GetString(,1, "����", "", "")
			Sql = Split(Sql,"����")
		End If
		Rs.Close : Set Rs = Nothing
		ToolsInfo = Sql
		ToolsSetting = Split(ToolsInfo(16),",")
	End Sub
	'---------------------------------------------------
	'��ȡ�û�������Ϣ
	'---------------------------------------------------
	Public Sub GetUserToolsInfo(G_USerID,G_ToolsID)
		Dim Sql,Rs
		G_USerID = CheckNumeric(G_USerID)
		G_ToolsID = CheckNumeric(G_ToolsID)
		If G_USerID = 0 or G_ToolsID = 0 Then ShowErr(3):Exit Sub
		'ID=0 ,UserID=1 ,UserName=2 ,ToolsID=3 ,ToolsName=4 ,ToolsCount=5 ,SaleCount=6 ,UpdateTime=7 ,SaleMoney=8 ,SaleTicket=9
		Sql = "Select ID,UserID,UserName,ToolsID,ToolsName,ToolsCount,SaleCount,UpdateTime,SaleMoney,SaleTicket From [Dv_Plus_Tools_buss] Where ToolsCount>0 and UserID="& G_USerID &" and ToolsID="& G_ToolsID
		Set Rs = Dvbbs.Plus_Execute(Sql)
		If Rs.Eof Then
			ShowErr(3):Exit Sub
		Else
			Sql = Rs.GetString(,1, "����", "", "")
			Sql = Split(Sql,"����")
		End If
		Rs.Close : Set Rs = Nothing
		UserToolsInfo = Sql
	End Sub
	'---------------------------------------------------
	'��ȡĿ���û���Ϣ
	'---------------------------------------------------
	Public Sub GetToUserInfo(ToUserID)
		Dim Sql,Rs
'		ToUserID = ToUserID
		If ToUserID = 0 Then ShowErr(11):Exit Sub
		'UserID=0,UserName=1,LockUser=2,UserPost=3,UserTopic=4,UserMoney=5,UserTicket=6,userWealth=7,userEP=8,userCP=9,UserPower=10,UserGroupID=11
		Sql = "Select UserID,UserName,LockUser,UserPost,UserTopic,UserMoney,UserTicket,userWealth,userEP,userCP,UserPower,UserGroupID From [Dv_User] Where UserID="& ToUserID
		Set Rs = Dvbbs.Execute(Sql)
		If Rs.Eof Then
			ShowErr(11):Exit Sub
		Else
			Sql = Rs.GetString(,1, "����", "", "")
			Sql = Split(Sql,"����")
		End If
		Rs.Close : Set Rs = Nothing
		ToUserInfo = Sql
	End Sub
	'---------------------------------------------------
	'����û�ʹ�õ���Ȩ��
	'---------------------------------------------------
	Public Sub ChkUseTools()
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		ChkUserGroup
		If Dvbbs.boardID>0 Then Chkboard
		If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text)<=cCur(ToolsInfo(7)) or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text)<=cCur(ToolsInfo(8)) or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text)<=cCur(ToolsInfo(9)) or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text)<=cCur(ToolsInfo(10))Then ShowErr(12):Exit Sub
		Call GetUserToolsInfo(Dvbbs.UserID,ToolsID)
	End Sub

	'---------------------------------------------------
	'���Ŀ���û�ʹ�õ���Ȩ��
	'---------------------------------------------------
	Public Sub ChkToUseTools(ToUserID)
		If Not IsArray(ToUserInfo) Then GetToUserInfo(ToUserID)
		If cCur(ToUserInfo(3))<=cCur(ToolsSetting(0)) or cCur(ToUserInfo(7))<=cCur(ToolsSetting(1)) or cCur(ToUserInfo(8))<=cCur(ToolsSetting(2)) or cCur(ToUserInfo(9))<=cCur(ToolsSetting(3)) Then ShowErr(13):Exit Sub
	End Sub

	'---------------------------------------------------
	'����û�������ʹ�õ���Ȩ��
	'---------------------------------------------------
	Public Sub ChkUserGroup()
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		If Cint(ToolsInfo(3)) = 0 Then ShowErr(6):Exit Sub
		If ToolsInfo(11) = "" or Instr(","& ToolsInfo(11) &",",","& Dvbbs.UserGroupID &",") = 0 Then ShowErr(4):Exit Sub
	End Sub
	'---------------------------------------------------
	'���������ʹ�õ���Ȩ��
	'---------------------------------------------------
	Public Sub Chkboard()
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		If ToolsInfo(12) = "" or Instr(","& ToolsInfo(12) &",",","& Dvbbs.boardID &",") = 0 Then ShowErr(5):Exit Sub
	End Sub
	
	Public Property Let buySum(byVal Value)
		buyCount = Value
	End Property

	'---------------------------------------------------
	'����û��������Ȩ�ޣ� bType �����ͣ�Ϊ�û�ѡȡ�Ĺ�������
	'---------------------------------------------------
	Public Sub ChkbuyTools(byval bType)
		Dim CanbuyTools
		CanbuyTools = False
		If bType="" or Not Isnumeric(bType) Then
			bType = -1
		Else
			bType = Cint(bType)
		End If
		If Not IsArray(ToolsInfo) Then GetToolsInfo
		If Int(ToolsInfo(4)) = 0 OR buyCount>Int(ToolsInfo(4)) OR buyCount = 0 Then ShowErr(8):Exit Sub '��治��
		Select Case Cint(ToolsInfo(14))
			Case 0 'ֻ����
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)>=Int(ToolsInfo(6))*buyCount and bType=0 Then
					CanbuyTools = True
				End If
			Case 1 'ֻ���ȯ
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)>=Int(ToolsInfo(13))*buyCount and bType=1 Then
					CanbuyTools = True
				End If
			Case 2 '���+��ȯ
				If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)<Int(ToolsInfo(6))*buyCount Or cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)<Int(ToolsInfo(13))*buyCount Then
					CanbuyTools = False
				Else
					CanbuyTools = True
				End If
			Case 3 '��һ��ȯ
				If bType=0 Then
					If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text)>Int(ToolsInfo(6))*buyCount Then CanbuyTools = True
				ElseIf bType=1 Then
					If cCur(Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text)>Int(ToolsInfo(13))*buyCount Then CanbuyTools = True
				Else
					CanbuyTools = False
				End If
			Case Else
				ShowErr(10):Exit Sub
		End Select
		If CanbuyTools = False Then ShowErr(7):Exit Sub
	End Sub
	'---------------------------------------------------
	'����ʽ
	'---------------------------------------------------
	Public Property Get buyType(byval bType)
		Select Case Cint(bType)
			Case 0 : buyType = "ֻ����"
			Case 1 : buyType = "ֻ���ȯ"
			Case 2 : buyType = "���+��ȯ"
			Case 3 : buyType = "��һ��ȯ"
			Case Else : buyType = "��ͣ����"
		End Select
		buyType = "<font class=redfont>"&buyType&"</font>"
	End Property

	Public Sub ShowErr(byval Code)
		If Code<>"" Then Response.redirect "showerr.asp?ErrCodes="& ErrCodes(Code) &"&action=NoHeadErr"
	End Sub
	'---------------------------------------------------
	'������Ϣ
	'---------------------------------------------------
	Public Function ErrCodes(byval ErrNum)
		Select Case ErrNum
			Case 1 : ErrCodes = "<li>���������Ѿ��رգ�</li>"
			Case 2 : ErrCodes = "<li>���߽��������Ѿ��رգ����ܽ��е��߽��ף�</li>"
			Case 3 : ErrCodes = "<li>�õ��߲����ڻ��������ȷ��</li>"
			Case 4 : ErrCodes = "<li>��û�й����ʹ�øõ��ߵ�Ȩ�ޣ�</li>"
			Case 5 : ErrCodes = "<li>����鲻��ʹ�øõ��ߣ�</li>"
			Case 6 : ErrCodes = "<li>�õ����ѱ�ϵͳ��ֹʹ�ã�</li>"
			Case 7 : ErrCodes = "<li>���Ľ�һ��ȯ�����ѡȡ�Ĺ���ʽ����ȷ�����ܹ���õ��ߣ�</li>"
			Case 8 : ErrCodes = "<li>�õ���ϵͳ��治�㣬��ͣ����</li>"
			Case 9 : ErrCodes = "<li>ת�õ������ѳ�������ӵ�еĵ������ݻ�û����д��ȷ�ĵ���������������ֹ��</li>"
			Case 10 : ErrCodes = "<li>��ͣ����</li>"
			Case 11 : ErrCodes = "<li>����ʹ��Ŀ���û������ڻ��������ȷ��"
			Case 12 : ErrCodes = "<li>����������������Ǯֵ�����ֵ������ֵ���㣬����û��ʹ�øõ��ߵ�Ȩ�ޣ�</li>"
			Case 13 : ErrCodes = "<li>����ʹ�õ�Ŀ���û������������Ǯֵ�����ֵ������ֵ���㣬�����㲻��ʹ�øõ��ߣ�</li>"
			Case 14 : ErrCodes = "<li>�˲�����������ͬ�û���֮����У�</li>"
			Case 15 : ErrCodes = "<li>���ҩֻ�������Լ������������ϣ�</li>"
			Case 16 : ErrCodes = "<li>�����õ�ת�ý�һ��ȯ������ȷ��</li>"
			Case 17 : ErrCodes = "<li>���Ľ�һ��ȯ���㣬����ת�ã�</li>"
			Case 18 : ErrCodes = "<li>���û�û���κε��ߡ�</li>"
		End Select
	End Function

	Public Function CheckNumeric(byval CHECK_ID)
		If CHECK_ID<>"" and IsNumeric(CHECK_ID) Then _
			CHECK_ID = Int(CHECK_ID) _
		Else _
			CHECK_ID = 0
		CheckNumeric = CHECK_ID
	End Function

End Class

'--------------------------------------------------------------------------------
'�û���Ϣ
'--------------------------------------------------------------------------------
Sub UserInfo()
	Dim Sql,Rs,UserToolsCount
	Sql = "Select Sum(ToolsCount) From [Dv_Plus_Tools_buss] where UserID="& Dvbbs.UserID
	Set Rs = Dvbbs.Plus_Execute(Sql)
	UserToolsCount = Rs(0)
	If IsNull(UserToolsCount) Then UserToolsCount = 0
%>
<table border="0" cellpadding="3" cellspacing="1" align="center" class="tableborder1" style="Width:100%">
	<tr>
		<th>��������</th>
	</tr>
	<tr>
		<td align="center" class="tablebody1">
			<table border="0" cellpadding="3" cellspacing="1" align="center" style="Width:90%">
				<tr>
					<td class="tablebody2" style="text-align:left">��ң�<b>
						<font color="<%=Dvbbs.mainsetting(1)%>">
							<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usermoney").text%>
						</font></b> ��
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">��ȯ��<b>
						<font color="<%=Dvbbs.mainsetting(1)%>">
							<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userticket").text%>
						</font></b> ��
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left">���ߣ�
						<a href="?action=UserTools_List"><b>
							<font color="<%=Dvbbs.mainsetting(1)%>"><%=UserToolsCount%></font></b></a> ��
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">��Ǯ��<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userwealth").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left">���£�<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpost").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">���飺<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userep").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody2" style="text-align:left">������<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usercp").text%>
					</td>
				</tr>
				<tr>
					<td class="tablebody1" style="text-align:left">������<%=Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@userpower").text%>
					</td>
				</tr>
				<tr><td class="tablebody2"></td></tr>
			</table>
		</td>
	</tr>
</table>
<%
End Sub

Sub Tools_Nav_Link()
%>
	<table border="0" width="<%=Dvbbs.mainsetting(0)%>" cellpadding="2" cellspacing="0" align="center">
		<tr>
		<th><a href="plus_Tools_Center.asp">ϵͳ��������</a></th>
		<th><a href="plus_Tools_Center.asp?action=UserbussTools_List" >�û���������</a></th>
		<th ><a href="?action=UserTools_List">�ҵĵ�����</a></th>
		<th><a href="UserPay.asp">������̳��ȯ</a></th>
		</tr>
	</table>
<%
End Sub
%>