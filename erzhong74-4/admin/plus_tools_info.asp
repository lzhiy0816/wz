<!--#include file =../conn.asp-->
<!--#include file="inc/const.asp"-->
<!--#include file="../inc/dv_clsother.asp"-->
<!--#include file="../Dv_plus/Tools/plus_Tools_const.asp"-->
<%
Head()
Dim admin_flag
If Trim(Request("action"))="Setting" Then Tools_Setting()
admin_flag=",40,"
CheckAdmin(admin_flag)
Main_head()
Select Case Trim(Request("action"))
	'Case "Setting" : Tools_Setting
	Case "AllStarSetting" : AllStarSetting
	Case "Editinfo" : AddTools
	Case "SaveTools" : SaveTools
	Case "List" : Tools_List()
	Case "UpdateUserStock" : UpdateUserStock
	Case Else
		Tools_List()
End Select
If founderr then call dvbbs_error()
footer()

'��������
Sub Main_head()
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<tr><th>�������Ĺ���</th></tr>
<tr><td class="td2"><B>������������˵��</B>��<BR>
	1��<font color=red>�������Ա�������ӿ�棬�밴��̳�г�ʵ��������е��ڣ�������Ƶ�����ӡ�</font>��ͨ���ڵ�����ӵ�����ǳ��ٵ�����½������ӿ�棬��ֹ��Ա�û�̧�߼۸񡣣�
	<BR>
	2��<font color=blue>ϵͳ�������Ϊ-1���ʾ�õ���Ϊϵͳ����</font>��ϵͳ�����ۣ�һ��Ϊ��̳ʹ�ù��̲���������û��õ����Լ�ת�û����</td></tr>
</table>
<br>
<%
End Sub

Sub Tools_Setting()
	Response.Redirect "setting.asp#settingxu"
End Sub

'�����б�
Sub Tools_List()
	Dim Orders
	Select case Trim(Request.QueryString("orders"))
		Case "0" : Orders = "SysStock"
		Case "1" : Orders = "UserMoney"
		Case "2" : Orders = "UserStock"
		Case "3" : Orders = "IsStar"
		Case Else : Orders = "SysStock"
	End Select
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr><th colspan="6" style="text-align:center;">���������б�</th></tr>
	<tr><td class=td1 colspan="6">
	[<a href="?action=UpdateUserStock"><font color=blue>�����û����</font></a>]
	</td></tr>
	<tr><td class="td1" colspan="6"><ol type="A">
	<li>����������ƽ����޸�����ϸ���ϡ�</li>
	<li>��������ӵı�����Ŀ�鿴��Ӧ������</li>
	</ol></td></tr>
	<tr>
		<th width="17%">����</th>
		<th width="45%">˵��</th>
		<th width="15%" id=TableTitleLink><A HREF="?orders=1" Title="��������ٵ�������">�۸�</A></th>
		<th width="5%" id=TableTitleLink><A HREF="?orders=0" Title="��ϵͳ������ٵ�������">���</A></th>
		<th width="10%" id=TableTitleLink><A HREF="?orders=2" Title="���û�ӵ�п�����ٵ�������">�û����</A></th>
		<th width="8%" id=TableTitleLink><A HREF="?orders=3" Title="���رյ���������">����</A></th>
	</tr>
	<form action="?action=AllStarSetting" method=post> 
	<%
	Dim Rs,Sql,i
	Sql = "Select ID,ToolsName,ToolsInfo,IsStar,SysStock,UserStock,UserMoney,UserTicket,BuyType From [Dv_Plus_Tools_Info] ORDER BY "& Orders
	Set Rs = Dvbbs.Plus_Execute(Sql)
	If Not Rs.eof Then
		SQL=Rs.GetRows(-1)
	Else
		Response.Write "<tr><td class=""td2"" colspan=""6"" align=center>���߻�δ���ӣ�</td></tr></form></table>"
		Exit Sub
	End If
	Rs.close:Set Rs = Nothing
	For i=0 To Ubound(SQL,2)
	%>
	<tr>
		<td class="td2"><a href="?action=Editinfo&EditID=<%=SQL(0,i)%>"><%=Server.Htmlencode(SQL(1,i))%></a></td>
		<td class="td2"><%=Server.Htmlencode(SQL(2,i)&"")%></td>
		<td class="td2">
		<%
		Select Case SQL(8,i)
		Case 0
			Response.Write SQL(6,i)
		Case 1
			Response.Write SQL(7,i)
		Case 2
			Response.Write SQL(6,i) & " And " & SQL(7,i)
		Case 3
			Response.Write SQL(6,i) & " Or " & SQL(7,i)
		End Select
		%>
		</td>
		<td class="td2" align=center>
		<%
		If SQL(4,i)="-1" Then
		Response.Write "ϵͳ"
		Else
		Response.Write SQL(4,i)
		End If
		%>
		</td>
		<td class="td2" align=center><%=SQL(5,i)%></td>
		<td class="td2" align=center><INPUT TYPE="checkbox" class="checkbox" NAME="Star" <%If SQL(3,i)="1" Then Response.Write "checked"%> value="<%=SQL(0,i)%>">
		</td>
	</tr>
	<%
	Next
	Response.Write "<tr><td class=""td1"" colspan=""6"" align=right><INPUT TYPE=""submit"" class=""button"" value=""�����޸�����""></td></tr></form></table>"
End Sub

'�����޸ĵ��߿�����ر�����
Sub AllStarSetting()
	Dim EditID,Sql
	EditID = Trim(Request.Form("Star"))
	If CheckID(EditID) = False Then
		Response.Write "�޸���ֹ����ȷ���ύ�Ĳ����Ƿ���ȷ�������ύ��<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a>"
		Exit Sub
	End If
	'�Ƚ��������û�ԭΪ�ر�״̬
	Sql = "Update [Dv_Plus_Tools_Info] Set IsStar=0"
	Dvbbs.Plus_Execute(Sql)

	'������ѡ��
	Sql = "Update [Dv_Plus_Tools_Info] Set IsStar=1 Where ID in (" & EditID & ")"
	Dvbbs.Plus_Execute(Sql)

	Dv_suc("�����޸ĵ��߿������óɹ���")
End Sub

'���ӣ��޸ĵ�����Ϣ
Sub AddTools()
Dim EditID,Rs,Sql
EditID = Trim(Request.QueryString("EditID"))
If EditID<>"" and IsNumeric(EditID) Then
	EditID = Cint(EditID)
Else
	Response.Write "�޸���ֹ����ȷ���ύ�Ĳ����Ƿ���ȷ�������ύ��<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a>"
	Exit Sub
End If

'ID=0 ,ToolsName=1 ,ToolsInfo=2 ,IsStar=3 ,SysStock=4 ,UserStock=5 ,UserMoney=6 ,UserPost=7 ,UserWealth=8 ,UserEp=9 ,UserCp=10 ,UserGroupID=11 ,BoardID=12,UserTicket=13,BuyType=14,ToolsImg=15,ToolsSetting=16
Dim ToolsSetting
Sql = "Select ID,ToolsName,ToolsInfo,IsStar,SysStock,UserStock,UserMoney,UserPost,UserWealth,UserEp,UserCp,UserGroupID,BoardID,UserTicket,BuyType,ToolsImg,ToolsSetting From [Dv_Plus_Tools_Info] Where ID="& EditID

Set Rs = Dvbbs.Plus_Execute(Sql)
If Rs.Eof Then
	Response.Write "���ҵĵ������ݲ����ڣ�<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a>"
	Exit Sub
Else
	Sql = Rs.GetString(,1, "����", "", "")
	Sql = Split(Sql,"����")
End If
Rs.Close
Set Rs = Nothing
ToolsSetting = Split(Sql(16),",")
If SQL(15)="" Then SQL(15)="Dv_plus/Tools/pic/None.jpg"
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
<form action="?action=SaveTools" method=post name=PlusTools onsubmit="GetData();">
<input type=hidden name="EditID" value="<%=SQL(0)%>">
<tr><th colspan="2" style="text-align:center;">������������</th></tr>
<tr><td colspan="2" height=23 class="td1" align=center>
<img src="../<%=Server.Htmlencode(SQL(15))%>" border=0></td></tr>
<tr>
<td class="td2" width="20%" align=right><b>��������</b></td>
<td class="td2" width="80%"><INPUT TYPE="text" NAME="ToolsName" value="<%=Server.Htmlencode(SQL(1))%>" size=35> ���ܳ���50���ַ���</td>
</tr>
<tr>
<td class="td2" align=right><b>���߼��</b></td>
<td class="td2">
<INPUT  NAME="ToolsInfo" TYPE="text" Style="width:80%" value="<%=Server.Htmlencode(SQL(2))%>"> ���ܳ���250���ַ���</td>
</tr>
<tr>
<td class="td2" align=right><b>����ͼ��</b></td>
<td class="td2">
<INPUT  NAME="ToolsImg" TYPE="text" Style="width:80%" value="<%=Server.Htmlencode(SQL(15))%>"> ����ȷ��дͼƬ·��</td>
</tr>
<tr>
<td class="td2" align=right><b>��ǰ����״̬</b></td>
<td class="td2">
<input type=radio class="radio" name="IsStar" value=0 <%If Sql(3)="0" Then%>checked<%End If%>>�ر�&nbsp;
<input type=radio class="radio" name="IsStar" value=1 <%If Sql(3)="1" Then%>checked<%End If%>>����&nbsp;
</td>
</tr>
<tr><th colspan="2" style="text-align:center;">���߽�������</th></tr>
<tr>
<td class="td2" align=right><b>��Ҽ۸�</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsMoney" value="<%=SQL(6)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>��ȯ�۸�</b></td>
<td class="td2"><INPUT TYPE="text" NAME="UserTicket" value="<%=SQL(13)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>��ǰϵͳ���</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsStock" value="<%=SQL(4)%>" size=10>&nbsp;��ֵΪ��-1��Ϊϵͳ���ߣ����������ס�</td>
</tr>
<tr>
<td class="td2" align=right><b>����ʽ</b></td>
<td class="td2">
<SELECT NAME="ToolsBuyType">
	<option value="0"<%If Sql(14)=0 Then%> Selected<%End If%>>ֻ����
	<option value="1"<%If Sql(14)=1 Then%> Selected<%End If%>>ֻ���ȯ
	<option value="2"<%If Sql(14)=2 Then%> Selected<%End If%>>���+��ȯ
	<option value="3"<%If Sql(14)=3 Then%> Selected<%End If%>>��һ��ȯ
</option>
</SELECT>
</td>
</tr>

<tr><th colspan="2" style="text-align:center;">����ʹ��Ȩ������</th></tr>
<tr>
<td class="td2" align=right><b>�û�����������</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsPost" value="<%=SQL(7)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>�û���Ǯ������</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsWealth" value="<%=SQL(8)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>�û�����ֵ����</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsEP" value="<%=SQL(9)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>�û�����ֵ����</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsCP" value="<%=SQL(10)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>ʹ��Ŀ���û�����������</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(0)" value="<%=ToolsSetting(0)%>" size=10> ���ֵ���ֻ��Ŀ���û�������������ʹ�ã���ͬ</td>
</tr>
<tr>
<td class="td2" align=right><b>ʹ��Ŀ���û���Ǯ������</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(1)" value="<%=ToolsSetting(1)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>ʹ��Ŀ���û�����ֵ����</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(2)" value="<%=ToolsSetting(2)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>ʹ��Ŀ���û�����ֵ����</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(3)" value="<%=ToolsSetting(3)%>" size=10></td>
</tr>
<tr>
<td class="td2" align=right><b>����ʹ�÷��㽱��</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsSetting(4)" value="<%=ToolsSetting(4)%>" size=10><BR> ���ֵ��߿����ã����ú�ʹ�øõ��߿ɻ��һ����ҵĽ������Զ��û���Ч�ĵ���ÿ����һ�ʽ������ʹ���û�һ����ҵĽ�����<font color=blue>�ڽ�ҹ���������Ӧ�����У�������Ϊ�ٷֱȣ���0.3�����繺�������Ҫ10����ң��򷵻�10*0.3�Ľ�Ҹ�������</font>��<font color=red>�ڲ�˰����Ϊ��Ŀ���û����յ��ܽ�����ٷֱ�</font></td>
</tr>
<tr>
<td class="td2" align=right><b>����ʹ�õ��û���ID</b></td>
<td class="td2"><INPUT TYPE="text" NAME="ToolsGroupID" value="<%=Trim(SQL(11))%>" size=40 Disabled>
<INPUT TYPE="hidden" NAME="TrueToolsGroupID" value="<%=Trim(SQL(11))%>">
<input type="button" class="button" value="����" onclick="PlusOpen('../plus_Tools_InfoSetting.asp?orders=0&id=<%=SQL(0)%>',650,500)"></td>
</tr>
<tr>
<td class="td2" align=right><b>����ʹ�õİ��ID</b></td>
<td class="td2">
<INPUT TYPE="text" NAME="ToolsBoardID" value="<%=Trim(SQL(12))%>" size=40 Disabled>
<INPUT TYPE="hidden" NAME="TrueToolsBoardID" value="<%=Trim(SQL(12))%>">
<input type="button" class="button" value="����" onclick="PlusOpen('../plus_Tools_InfoSetting.asp?orders=1&id=<%=SQL(0)%>',650,500)"></td>
</tr>
<tr><td class="td1" colspan="2" align=center><INPUT TYPE="submit" class="button" value="�����޸�����"></td></tr>
</form>
</table>
<SCRIPT LANGUAGE="JavaScript">
<!--
function GetData(){
PlusTools.TrueToolsGroupID.value=PlusTools.ToolsGroupID.value;
PlusTools.TrueToolsBoardID.value=PlusTools.ToolsBoardID.value;
}
//-->
</SCRIPT>
<%
End Sub

'�������ӣ��޸ĵ�����Ϣ
Sub SaveTools()
	Dim EditID,ToolsName,ToolsInfo,ToolsImg
	Dim ToolsGroupID,ToolsBoardID
	Dim Rs,Sql,i
	Dim ToolsSetting,ChanSetting
	EditID = CheckNumeric(Request.Form("EditID"))
	ToolsName = Left(Trim(Request.Form("ToolsName")),50)
	ToolsInfo = Left(Trim(Request.Form("ToolsInfo")),250)
	ToolsImg = Left(Trim(Request.Form("ToolsImg")),150)
	ToolsGroupID = Trim(Replace(Request.Form("TrueToolsGroupID")," ",""))
	ToolsBoardID = Trim(Replace(Request.Form("TrueToolsBoardID")," ",""))
	If Right(ToolsGroupID,1)="," Then ToolsGroupID = Left(ToolsGroupID,Len(ToolsGroupID)-1)
	If Right(ToolsBoardID,1)="," Then ToolsBoardID = Left(ToolsBoardID,Len(ToolsBoardID)-1)
	If EditID = 0 Then
		Response.Write "���µĵ������ݲ����ڣ�<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a>"
		Exit Sub
	End If
	If CheckID(ToolsGroupID)=False Then ToolsGroupID = ""
	If CheckID(ToolsBoardID)=False Then ToolsBoardID = ""

	If ToolsName="" or ToolsInfo="" Then
		Response.Write "�޸���ֹ���������ƻ��鲻��Ϊ�գ�<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a>"
		Exit Sub
	Else
		ToolsName = Replace(ToolsName,"����","")
		ToolsInfo = Replace(ToolsInfo,"����","")
	End If
	For i=0 To 60
		If Request.Form("ToolsSetting("&i&")")="" Then
			ChanSetting = 0
		Else
			ChanSetting = Replace(Request.Form("ToolsSetting("&i&")"),",","")
		End If
		If i = 0 Then
			ToolsSetting = ChanSetting
		Else
			ToolsSetting = ToolsSetting & "," & ChanSetting
		End If
	Next

	Set Rs = Dvbbs.iCreateObject("adodb.recordset")
	Sql = "Select * From [Dv_Plus_Tools_Info] where ID="& EditID
	If Cint(Dvbbs.Forum_Setting(92))=1 Then
		If Not IsObject(Plus_Conn) Then Plus_ConnectionDatabase
		Rs.Open Sql,Plus_Conn,1,3
	Else
		If Not IsObject(Conn) Then ConnectionDatabase
		Rs.Open Sql,conn,1,3
	End IF
	If Rs.eof and Rs.bof then
		Response.Write "���ҵĵ������ݲ����ڣ�<a href="&Request.ServerVariables("HTTP_REFERER")&" ><<������һҳ</a>"
		Exit Sub
	Else
		Rs("ToolsName") = ToolsName
		Rs("ToolsInfo") = ToolsInfo
		Rs("ToolsImg") = ToolsImg
		Rs("IsStar") = CheckNumeric(Request.Form("IsStar"))
		Rs("SysStock") = CheckNumeric(Request.Form("ToolsStock"))
		Rs("UserTicket") = CheckNumeric(Request.Form("UserTicket"))
		Rs("UserMoney") = CheckNumeric(Request.Form("ToolsMoney"))
		Rs("UserPost") = CheckNumeric(Request.Form("ToolsPost"))
		Rs("UserWealth") = CheckNumeric(Request.Form("ToolsWealth"))
		Rs("UserEp") = CheckNumeric(Request.Form("ToolsEP"))
		Rs("UserCp") = CheckNumeric(Request.Form("ToolsCP"))
		Rs("UserGroupID") = ToolsGroupID
		Rs("BoardID") = ToolsBoardID
		Rs("BuyType") = CheckNumeric(Request.Form("ToolsBuyType"))
		Rs("ToolsSetting") = ToolsSetting
		Rs.Update
	End If 
	Rs.Close
	Set Rs = Nothing
	Dvbbs.Plus_Execute("UPDATE [Dv_Plus_Tools_Buss] Set ToolsName ='"& Dvbbs.Checkstr(ToolsName) &"' where ToolsID="&EditID)
	Dv_suc(ToolsName&"���߿������óɹ���")
End Sub

'ɾ��������Ϣ
Sub DllTools()

End Sub

'���µ��ߵ��û�ӵ�п��
Sub UpdateUserStock()
	Dim Rs,Sql,Totals
	If IsSqlDataBase = 1 Then
		Sql = "Update [Dv_Plus_Tools_Info] Set UserStock = (Select Count(*) From [Dv_Plus_Tools_Buss] where ToolsID=Dv_Plus_Tools_Info.ID)"
		Dvbbs.Plus_Execute(Sql)
	Else
		Sql = "Select ID From [Dv_Plus_Tools_Info]"
		Set Rs = Dvbbs.Plus_Execute(Sql)
		Do while Not Rs.eof
			Totals = Dvbbs.Plus_Execute("Select Count(*) From [Dv_Plus_Tools_Buss] where ToolsID="&Rs(0))(0)
			Dvbbs.Plus_Execute("Update [Dv_Plus_Tools_Info] Set UserStock = "& Totals &" where id="&Rs(0))
		Rs.movenext
		loop
		Rs.Close
		Set Rs = Nothing
	End If
	Dv_suc("�û�ӵ�п����³ɹ���")
End Sub

Function CheckID(CHECK_ID)
	Dim Fixid
	CheckID = False
	Fixid = Replace(CHECK_ID,",","")
	Fixid = Trim(Replace(Fixid," ",""))
	If IsNumeric(Fixid) and Fixid<>"" Then CheckID = True
End Function

Function CheckNumeric(CHECK_ID)
	If CHECK_ID<>"" and IsNumeric(CHECK_ID) Then
		CHECK_ID = cCur(CHECK_ID)
	Else
		CHECK_ID = 0
	End If
	CheckNumeric = CHECK_ID
End Function
%>