<!--#include file=../conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="../Dv_plus/IndivGroup/Dv_IndivGroup_Config.asp" -->
<%
Dvbbs.ShowSQL=1

Head()
dim admin_flag,info
admin_flag=",8,"
CheckAdmin(admin_flag)
Top_Nav()
Call main()
Footer()

Sub main()
	If Not IsObject(Dv_IndivGroup_Conn) Then Dv_IndivGroup_ConnectionDatabase
	Select Case Lcase(Request("action"))
		Case "grouplist" : GroupInfo()
		Case "save" : savegroup()
		Case "del" : del()
		Case "groupuserlist" : GroupUserList()
		Case "�����û�ͳ��" : upusernum()
		Case "manage" : UserManage()
		Case "groupclasslist":GroupClassList()
		Case "updategroupclass":UpdateGroupClass()
		Case "deletegroupclass":DeleteGroupClass()
		Case Else
			GroupInfo()
	End Select
End Sub

Sub Top_Nav()
%>
<script language="JavaScript" src="../inc/Pagination.js"></script>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<th width="100%" style="text-align:center;" colspan=2>��̳Ȧ�ӹ���</th>
	</tr>
	<tr>
		<td height="23" colspan="2" class=td1>���ڲ������Ĳ˵����������һ���˵�������Ȧ�ӡ���</td>
	</tr>
	<tr>
		<td height="23" colspan="2" class=td1><b>�û�Ȧ�ӹ���</b>������������޸Ļ���ɾ����̳Ȧ�ӡ�</td>
	</tr>
	<tr>
		<td height="23" colspan="2" class=td1>
			<button onclick="window.location='indivGroup.asp?action=addedit&orders=0'" class="button">���Ȧ��</button>&nbsp;
			<button onclick="window.location='indivGroup.asp'" class="button">Ȧ�ӹ���</button>&nbsp;
			<button onclick="window.location='indivGroup.asp?action=groupuserlist'" class="button">Ȧ���û�����</button>&nbsp;
			<button onclick="window.location='indivGroup.asp?action=groupclasslist'" class="button">Ȧ�ӷ������</button>&nbsp;
		</td>
	</tr>
</table>
<br/>
<%
End Sub

'Ȧ���б�
Sub GroupInfo()
	Dim Rs,SQL,GroupID
	Dim Page,Orders,ordername,keyword,SQLStr,QueryString
	Dim TotalRec,i,Pcount
	Dim AppUserID,UserGroup,Locked,ViewFlag
	Dim GroupName,GroupInfo,AppUserName,GroupStats,LimitUser

	If LCase(Request("action"))="pass" Then
		GroupID=Dvbbs.CheckNumeric(Request("GroupID"))
		Dv_IndivGroup_Conn.Execute("Update Dv_GroupName Set Stats=1,UserNum=1,PassDate='"&now()&"' Where ID="&GroupID)
		Set Rs=Dv_IndivGroup_Conn.Execute("Select ID,AppUserID,AppUserName From Dv_GroupName Where ID="&GroupID&" Order By ID Desc")
		If Not Rs.Eof Then
			Dv_IndivGroup_Conn.Execute("Insert Into Dv_GroupUser(GroupID,UserID,UserName,IsLock) Values("&Rs(0)&","&Rs(1)&",'"&Rs(2)&"',2)")
			AppUserID = Dvbbs.CheckNumeric(Rs(1))
			Rs.Close
			Set Rs=Dvbbs.Execute("Select UserGroup From Dv_User Where UserID="&AppUserID)
			If Not Rs.Eof Then
				UserGroup=Rs(0)
				If IsNull(UserGroup) Or UserGroup="" Then
					UserGroup = ","&GroupID&","
				Else
					If InStr(UserGroup,","&GroupID&",")<1 Then UserGroup = UserGroup&GroupID&","
				End If
				Dvbbs.Execute("Update Dv_User Set UserGroup='"&UserGroup&"' Where UserID="&AppUserID)
			End If
		End If
		Rs.Close:Set Rs=Nothing
	End If

	If LCase(Request("action"))="batchmodify" Then
		If Request.Form("groupid")="" Then
			Response.write "<script language='javascript'>alert('��ѡ��Ҫ�޸ĵ���');</script>"
		Else
			For i=1 To Request.Form("groupid").Count
				GroupID = Dvbbs.CheckNumeric(Request.Form("groupid")(i))
				LimitUser = Dvbbs.CheckNumeric(Request.Form("limituser_"&GroupID))
				GroupStats = Dvbbs.CheckNumeric(Request.Form("GroupStats_"&GroupID))
				Set Rs = Dv_IndivGroup_Conn.Execute("Select ID,GroupName,GroupInfo,AppUserID,AppUserName,Stats,LimitUser,Locked,ViewFlag From Dv_GroupName Where ID="&GroupID)
				If Not Rs.Eof Then
					SQL = Rs.GetRows(1)
					Rs.Close
					If SQL(5,0)=0 And SQL(5,0) < GroupStats Then
						Set Rs = Dv_IndivGroup_Conn.Execute("Select * From Dv_GroupUser Where GroupID="&GroupID&" And UserID="&SQL(3,0))
						If Rs.Eof Then Dv_IndivGroup_Conn.Execute("Insert Into Dv_GroupUser(GroupID,UserID,UserName,IsLock) Values("&SQL(0,0)&","&SQL(3,0)&",'"&SQL(4,0)&"',2)")
						Rs.Close

						Set Rs=Dvbbs.Execute("Select UserGroup From Dv_User Where UserID="&SQL(3,0))
						If Not Rs.Eof Then
							UserGroup=Rs(0)
							If IsNull(UserGroup) Or UserGroup="" Then
								UserGroup = ","&GroupID&","
							Else
								If InStr(UserGroup,","&GroupID&",")<1 Then UserGroup = UserGroup&GroupID&","
							End If
							Dvbbs.Execute("Update Dv_User Set UserGroup='"&UserGroup&"' Where UserID="&SQL(3,0))
						End If
					End If
					Dv_IndivGroup_Conn.Execute("Update Dv_GroupName Set LimitUser="&LimitUser&",Stats="&GroupStats&" Where ID="&GroupID)
				End If
			Next
			Response.write "<script language='javascript'>alert('�޸ĳɹ���');</script>"
		End If
	End If

	TotalRec=0
	Page=Dvbbs.CheckNumeric(Request("page")):If Page=0 Then Page=1
	QueryString = "page="&page
	Orders=Dvbbs.CheckNumeric(Request("orders"))
	QueryString = QueryString & "&orders="&orders
	keyword=Dvbbs.CheckStr(Request("keyword"))
	If keyword<>"" Then	QueryString = QueryString & "&keyword="&keyword

	SQL="id,groupname,groupinfo,appuserid,appusername,usernum,stats,postnum,topicnum,limituser,AppDate,PassDate,Locked,ViewFlag"
	If orders=1 Then
		SQLStr = "Where Stats=0"
	ElseIf orders=0 Then
		SQLStr = ""
	Else
		SQLStr = "Where Stats>0"
	End If
	If keyword<>"" Then
		If SQLStr="" Then
			SQLStr = "Where GroupName like '%"&keyword&"%'"
		Else
			SQLStr = SQLStr&" And GroupName like '%"&keyword&"%'"
		End If
	End if
	Select Case orders
		Case 0
			SQL="Select "&SQL&" From [Dv_GroupName] "&SQLStr&" Order By ID Desc"
		Case 1
			SQL="Select "&SQL&" From [Dv_GroupName] "&SQLStr&" Order By ID Desc"
		Case 2
			SQL="Select "&SQL&" From [Dv_GroupName] "&SQLStr&" Order By ID Desc"
		Case 3
			SQL="Select top 20 "&SQL&" From [Dv_GroupName] "&SQLStr&" Order By PostNum Desc"
		Case 4
			SQL="Select top 20 "&SQL&" From [Dv_GroupName] "&SQLStr&" Order By UserNum Desc"
		Case Else
			SQL="Select "&SQL&" From [Dv_GroupName] "&SQLStr&" Order By ID Desc"
	End Select

	If LCase(Request("action"))="addedit" Then
		GroupID=Dvbbs.CheckNumeric(Request("GroupID"))
		Set Rs=Dv_IndivGroup_Conn.Execute("Select GroupName,GroupInfo,AppUserName,Stats,LimitUser,Locked,ViewFlag From Dv_GroupName Where id="&GroupID)
		If Not Rs.Eof Then
			GroupName = Rs("GroupName")
			GroupInfo = Rs("GroupInfo")
			AppUserName = Rs("AppUserName")
			GroupStats  =Rs("Stats")
			LimitUser = Rs("LimitUser")
			Locked = Rs("Locked")
			ViewFlag = Rs("ViewFlag")
		Else
			GroupName = ""
			GroupInfo = ""
			AppUserName = ""
			GroupStats  = 1
			LimitUser = ""
			Locked = 0
			ViewFlag = 1
		End If
		Rs.Close:Set Rs=Nothing
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<form method="POST" action="indivGroup.asp?action=save&<%=QueryString%>" name="groupform">
	<input type="hidden" name="groupid" value="<%=GroupID%>">
	<tr>
		<th width="100%" style="text-align:center;" colspan=2><b>���/�޸�Ȧ��</b></th>
	</tr>
	<tr>
		<td width="20%" class=td1><b>Ȧ������</b></td>
		<td width="80%" class=td1>
		<input type="text" name="Groupname" size="35" value="<%=Groupname%>">
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>����Ա����</b></td>
		<td width="80%" class=td1>
		<input type="text" name="AppUserName" size="35" value="<%=AppUserName%>">
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>Ȧ��˵��</b></td>
		<td width="80%" class=td1>
		<textarea rows="4" cols="35" name="GroupInfo"><%=GroupInfo%></textarea>
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>��Ա��������</b></td>
		<td width="80%" class=td1>
		<input type="text" name="LimitUser" size="35" value="<%=LimitUser%>">
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>Ȧ��״̬</b></td>
		<td width="80%" class=td1>
		<%If GroupStats=0 Then Response.write "<input type=""radio"" class=""radio"" name=""GroupStats"" value=""0""/>���"%>
		<input type="radio" class="radio" name="GroupStats" value="1"/>����
		<input type="radio" class="radio" name="GroupStats" value="2"/>����
		<input type="radio" class="radio" name="GroupStats" value="3"/>�ر�
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>��Ա����</b></td>
		<td width="80%" class=td1>
		<input type="radio" class="radio" name="locked" value="0"/>���ɼ���
		<input type="radio" class="radio" name="locked" value="1"/>��Ҫ���
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>�������</b></td>
		<td width="80%" class=td1>
		<input type="radio" class="radio" name="viewflag" value="0"/>����
		<input type="radio" class="radio" name="viewflag" value="1"/>������
		</td>
	</tr>
	<tr><td colspan="2">���������κ��˶������������������ֻ�г�Ա�������</td></tr>
	<tr><td height="23" colspan="2" class=td1><input type="submit" class="button" name="Submit" value="�� ��"></tr>
	</form>
	<script language="JavaScript">
	<!--
	chkradio(document.groupform.GroupStats,'<%=GroupStats%>')
	chkradio(document.groupform.locked,'<%=Locked%>')
	chkradio(document.groupform.viewflag,'<%=viewflag%>')
	//-->
	</script>
</table>
<br />
<%
	End If
%>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<form method="post" action="indivGroup.asp" name="searchform">
	<tr>
		<td colspan="10">
		&nbsp;����Ȧ��������<input name="keyword" type="text" value="<%=keyword%>" size="15" /> <input name="Submit" type="submit" class="button" value="����" />
		<select name="orders" onchange='javascript:submit()'>
			<option value="0" selected>����Ȧ��</option>
			<option value="1">�����Ȧ��</option>
			<option value="2">����Ȧ��</option>
			<option value="3">��������Top20</option>
			<option value="4">��Ա����Top20</option>
		</select>
		<script language="javascript">
		function setorders(ordernum){
			document.searchform.orders.options(ordernum).selected=1;
		}
		setorders(<%=orders%>);
		</script>

		</td>
	</tr>
	</form>
	<form method="post" action="indivGroup.asp?action=batchmodify&<%=QueryString%>" name="batchform">
	<tr>
		<th style="text-align:center;">&nbsp;</th>
		<th style="text-align:center;">Ȧ������</th>
		<th style="text-align:center;">Ȧ������������</th>
		<th style="text-align:center;">�û�����</th>
		<th style="text-align:center;">��������</th>
		<th style="text-align:center;">��������</th>
		<th style="text-align:center;">��Ա��������</th>
		<th style="text-align:center;">����ʱ��</th>
		<th style="text-align:center;">״̬</th>
		<th style="text-align:center;">����</th>
	</tr>
<%
	Const EachPageCount=20
	Set Rs=Dvbbs.iCreateObject("ADODB.RecordSet")
	Rs.Open SQL,Dv_IndivGroup_Conn,1,1
	Dvbbs.SqlQueryNum = Dvbbs.SqlQueryNum + 1
	If Not Rs.Eof Then
		TotalRec=Rs.RecordCount
		If TotalRec Mod EachPageCount=0 Then
			Pcount= TotalRec \ EachPageCount
		Else
			Pcount= TotalRec \ EachPageCount+1
		End If

		If orders=3 Or orders=4 Then
			SQL=Rs.GetRows(-1)
		Else
			RS.MoveFirst
			If Page > Pcount Then Page = Pcount
			If Page < 1 Then Page=1
			RS.Move (Page-1) * EachPageCount
			SQL=Rs.GetRows(EachPageCount)
		End If
		Rs.Close:Set Rs=Nothing
		'id=0,groupname=1,groupinfo=2,appuserid=3,appusername=4,usernum=5,stats=6,postnum=7,topicnum=8,limituser=9,appDate=10,PassDate=11
		For i = 0 To Ubound(SQL,2)
%>
	<tr>
		<td class="td1"><input type="checkbox" name="groupid" value="<%=SQL(0,i)%>"/></td>
		<td class="td1"><a title="<%=Dvbbs.Replacehtml(SQL(2,i))%>"><%=SQL(1,i)%></a></td>
		<td class="td1"><a href="../dispuser.asp?id=<%=SQL(3,i)%>" target="_blank"><%=SQL(4,i)%></a></td>
		<td class="td1"><%=SQL(5,i)%></td>
		<td class="td1"><%=SQL(7,i)%></td>
		<td class="td1"><%=SQL(8,i)%></td>
		<td class="td1" align="center"><input type="text" name="limituser_<%=SQL(0,i)%>" value="<%=SQL(9,i)%>" size="8"/></td>
		<td class="td1"><%=SQL(10,i)%></td>
		<td class="td1">
		<select size="1" name="GroupStats_<%=SQL(0,i)%>">
			<option value="1">����</option>
			<option value="2">����</option>
			<option value="3">�ر�</option>
			<option value="0">���</option>
		</select>
		<script language="javascript">ChkSelected(document.batchform.GroupStats_<%=SQL(0,i)%>,'<%=SQL(6,i)%>');</script>
		<%=GroupStatsStr(SQL(6,i))%>
		</td>
		<td height="23" class="td2" align="right">
			<%if SQL(6,i)=0 Then Response.write "<a href=""indivGroup.asp?groupid="&SQL(0,i)&"&action=pass&"&QueryString&""">ͨ��</a> | "%>
			<a href="indivGroup.asp?groupid=<%=SQL(0,i)%>&action=addedit&<%=QueryString%>">�޸�</a> |
			<a href="indivGroup.asp?groupid=<%=SQL(0,i)%>&action=del&<%=QueryString%>">ɾ��</a> |
			<a href="indivGroup.asp?action=GroupUserList&groupid=<%=SQL(0,i)%>">�鿴�û�</a>
		</td>
	</tr>
<%
		Next
	Else
		If orders=1 Then
			Response.write "<tr><td class=""td1"" colspan=""10"">û�д���˵�Ȧ��</td></tr>"
		Else
			Response.write "<tr><td class=""td1"" colspan=""10"">û�ҵ���������������Ȧ��</td></tr>"
		End If
		Rs.Close:Set Rs=Nothing
	End If
%>
	<tr>
		<td colspan="10">
			<div style="margin-left:35px;">
			<input type="submit" name="Submit" value="�����޸�" onclick="return clicksubmit(this.form);" class="button"/>
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);" class="chkbox" />&nbsp;ȫѡ/ȡ��
			</div>
		</td>
	</tr>
	</form>
</table>
<script language="javascript">
PageList('<%=page%>',10,'<%=EachPageCount%>','<%=TotalRec%>','?action=grouplist',1);
function CheckAll(form){
	for (var i=0;i<form.elements.length;i++){
		var e = form.elements[i];
		if (e.name == 'groupid') e.checked = form.chkall.checked;
	}
}
function clicksubmit(form){
	if(confirm('��ȷ��ִ�еĲ�����?')){
		this.document.batchform.submit();
		return true;
	}else{
		return false;
	}
}
</script>
<%
End Sub

Sub GroupClassList() Rem Ȧ�ӷ���
	Dim Sql,SqlStr,Rs
	Set Rs = Dvbbs.Execute("Select ID,ClassName,GroupCount,Orders From Dv_Group_Class Order By Orders,Id")
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
		<form method="post" action="indivGroup.asp?action=UpdateGroupClass" name="searchform">
		<tr>
			<td width="5%" style="text-align:center;">ѡ��</td>
			<td width="5%" style="text-align:center;">ID</td>
			<td style="text-align:center;">Ȧ�ӷ�������</td>
			<td width="10%" style="text-align:center;">���� ��</td>
			<td width="10%" style="text-align:center;">Ȧ������</td>
			<td width="10%" style="text-align:center;">�������</td>
		</tr>
		<%Do While Not Rs.EOF%>
		<tr>
			<td class="td1" style="text-align:center;"><input type="hidden" name="Id" value="<%=Rs(0)%>"/>&nbsp;</td>
			<td class="td1" style="text-align:center;"><%=Rs(0)%></td>
			<td class="td1"><input type="text" size="35" name="ClassName" value="<%=Rs(1)%>" /></td>
			<td class="td1"><input type="text" size="4" name="Orders" maxlength="3" value="<%=Rs(3)%>" /></td>
			<td class="td1" style="text-align:center;"><a href="../IndivGroup_List.asp?ClassId=<%=Rs(0)%>" target="_blank"><%=Rs(2)%></a></td>
			<td width="10%" style="text-align:center;"><a href="indivGroup.asp?action=DeleteGroupClass&Id=<%=Rs(0)%>" onclick="{if(confirm('��ȷ��Ҫɾ����Ȧ�ӷ��� [<%=Rs(1)%>] �𣿣�ɾ���������ɻָ���')){return true;}return false;}">ɾ������</a></td>
		</tr>
		<%
		Rs.MoveNext
		Loop
		%>
		<tr>
			<td class="td1" style="text-align:center;"><input type="checkbox" name="addClass" value="1"/></td>
			<td class="td1" style="text-align:center;"><font color="#FF0000">����</font></td>
			<td class="td1"><input type="text" size="35" name="addClassName" /></td>
			<td class="td1"><input type="text" size="4" name="addOrders" maxlength="3" value="0" /></td>
			<td class="td1" style="text-align:center;">&nbsp;</td>
			<td width="10%" style="text-align:center;">&nbsp;</td>
		</tr>
		<tr>
		<td colspan="6" style="text-align:center;">
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);" class="chkbox" />&nbsp;ȫѡ/ȡ��&nbsp;&nbsp;
			<input type="submit" name="Submit" value="�����޸�" onclick="" class="button"/>
		</td>
		</tr>
		</form>
	</table>
	<%
End Sub

Sub UpdateGroupClass()
	Dim Id,ClassName,Orders,i,Str,Rs
	For i=1 To Request.Form("Id").Count
		Id=Replace(Request.Form("Id")(i),"'","")
		ClassName=Replace(Request.Form("ClassName")(i),"'","")
		Orders=CInt(Replace(Request.Form("Orders")(i),"'",""))
		ClassName=Trim(ClassName)
		If IsNumeric(Id) And IsNumeric(Orders) And ClassName<>"" Then
			Dvbbs.Execute("Update Dv_Group_Class Set ClassName='"&ClassName&"',GroupCount=(Select Count(ClassId) From Dv_GroupName Where ClassId="&Id&"),Orders="&Orders&" Where Id="&Id)
		Else
			Str = "<font color=""#FF0000"">Id:"&Id&" Ȧ�����Ʋ���Ϊ�գ�������������֣�</font>"
			Exit For
		End If
	Next

	If CInt(Request.Form("addClass"))=1 Then
		ClassName = Trim(Replace(Request.Form("addClassName"),"'",""))
		Orders = Cint(Request.Form("addOrders"))
		Set Rs = Dvbbs.Execute("Select ID,ClassName From Dv_Group_Class Where ClassName='"&ClassName&"'")
		If Not Rs.Eof Then
			Str = "<br /><font color=""#FF0000"">Error: �Ѿ���������Ȧ�ӷ��������,��������д��</font>"
		Else
			If ClassName<>"" Then
				Dvbbs.Execute("Insert Into [Dv_Group_Class] (ClassName,GroupCount,Orders) Values('"&ClassName&"',0,"&Orders&")")
				Str = "<br /><font color=""#FF0000"">����Ȧ�ӷ��� <b>"&ClassName&"</b> �ɹ���</font>"
			Else
				Str = "<br /><font color=""#FF0000"">Error: ����Ȧ�ӷ���,����д��Ȧ�ӷ������ƣ�</font>"
			End If
		End If
	End If
	LoadGroupClass()
	Dv_suc("����Ȧ�ӷ������ݳɹ���"& Str)
End Sub

Sub DeleteGroupClass()
	Dim Id
	Id=CLng(Request("Id"))
	Dvbbs.Execute("Update Dv_GroupName Set ClassId=0 Where ClassId="&Id)
	Dvbbs.Execute("Delete From Dv_Group_Class Where Id="&Id)
	Dv_suc("ɾ��Ȧ�ӷ���ɹ���ԭ����Ȧ���Ѿ�ȡ�����ࡣ")
End Sub

'����Ȧ����Ϣ
Sub SaveGroup()
	Dim GroupID,GroupName,AppUserID,AppUserName,GroupInfo,GroupStats,LimitUser,Locked,ViewFlag
	Dim Rs,SQL,QueryStr
	Dim UserGroup

	GroupID = Dvbbs.CheckNumeric(Request("GroupID"))
	GroupName = Dvbbs.Checkstr(trim(Request("GroupName")))
	AppUserName = Dvbbs.Checkstr(trim(Request("AppUserName")))
	GroupInfo = Dvbbs.Checkstr(Request("GroupInfo"))
	GroupStats = Dvbbs.CheckNumeric(Request("GroupStats"))
	LimitUser = Dvbbs.CheckNumeric(Request("LimitUser"))
	Locked = Dvbbs.CheckNumeric(Request("Locked"))
	ViewFlag = Dvbbs.CheckNumeric(Request("ViewFlag"))
	Errmsg = ""
	If GroupID=0 Then
		If GroupName = "" Then Errmsg=ErrMsg + "<BR><li>���Ʋ���Ϊ�ա�"
		If AppUserName = "" Then Errmsg=ErrMsg + "<BR><li>Ȧ�ӹ���Ա����Ϊ�ա�"
		QueryStr = ""
	Else
		QueryStr = "And ID<>"&GroupID
	End If

	If GroupName<>"" Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select * From Dv_GroupName Where GroupName='"&GroupName&"' "&QueryStr&" Order By ID Desc")
		If Not Rs.Eof Then Errmsg=ErrMsg + "<BR><li>��Ȧ�������Ѿ������룬����д����Ȧ�����ơ�"
		Rs.Close:Set Rs=Nothing
	End If
	If AppUserName<>"" Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select * From Dv_GroupName Where AppUserName='"&AppUserName&"' "&QueryStr&" Order By ID Desc")
		If Not Rs.Eof Then Errmsg=ErrMsg + "<BR><li>����û��Ѿ�������Ȧ���ˣ�һ���û�ֻ������һ��Ȧ�ӡ�"
		Rs.Close:Set Rs=Nothing

		Set Rs=Dv_IndivGroup_Conn.Execute("Select UserID,UserGroupID,UserGroup From Dv_user Where UserName='"&AppUserName&"'")
		If Not Rs.Eof Then
			If Rs(1)<>5 Then
				AppUserID = Rs(0)
				UserGroup = Dvbbs.CheckStr(Rs(2))
			Else
				Errmsg=ErrMsg + "<BR><li>�û� "&AppUserName&" �ǵȴ���֤��(COPPA)��Ա��������ΪȦ�ӹ���Ա��"
			End If
		Else
			Errmsg=ErrMsg + "<BR><li>�û� "&AppUserName&" ��û��ע�ᣬ������ΪȦ�ӹ���Ա��"
		End If
		Rs.Close:Set Rs=Nothing
	End If

	If Errmsg<>"" Then
		dvbbs_error()
		Exit sub
	End If

	If GroupID=0 Then
		If LimitUser=0 Then LimitUser = 50
		If GroupStats>0 Then
			Dv_IndivGroup_Conn.Execute("Insert Into Dv_GroupName(GroupName,GroupInfo,AppUserID,AppUserName,UserNum,Stats,LimitUser,AppDate,PassDate,Locked,ViewFlag) Values('"&GroupName&"','"&GroupInfo&"',"&AppUserID&",'"&AppUserName&"',1,"&GroupStats&","&LimitUser&","&SqlNowString&","&SqlNowString&","&Locked&","&ViewFlag&")")
			Set Rs=Dv_IndivGroup_Conn.Execute("Select ID,AppUserID,AppUserName From Dv_GroupName Where GroupName='"&GroupName&"' Order By ID Desc")
			If Not Rs.Eof Then
				GroupID = Rs(0)
				Dv_IndivGroup_Conn.Execute("Insert Into Dv_GroupUser(GroupID,UserID,UserName,IsLock) Values("&Rs(0)&","&Rs(1)&",'"&Rs(2)&"',2)")
			End If
			Rs.Close:Set Rs=Nothing
		Else
			Dv_IndivGroup_Conn.Execute("Insert Into Dv_GroupName (GroupName,GroupInfo,AppUserID,AppUserName,UserNum,Stats,LimitUser,AppDate,PassDate,Locked,ViewFlag) Values ('"&GroupName&"','"&GroupInfo&"',"&AppUserID&",'"&AppUserName&"',1,"&GroupStats&","&LimitUser&","&SqlNowString&","&SqlNowString&","&Locked&","&ViewFlag&")")
		End If
		Dv_suc("<b>��ӳɹ���</b>")
	Else
		Set Rs=Dvbbs.iCreateObject("ADODB.RecordSet")
		SQL = "Select * From Dv_GroupName Where ID="&GroupID
		Rs.Open SQL,Dv_IndivGroup_Conn,1,3
		If Not Rs.Eof Then
			If GroupName<>"" Then Rs("GroupName")=GroupName
			If AppUserName<>"" Then
				Rs("AppUserID")=AppUserID
				Rs("AppUserName")=AppUserName
			End If
			If GroupInfo<>"" Then Rs("GroupInfo")=GroupInfo
			If LimitUser>0 Then Rs("LimitUser")=LimitUser
			If Not IsDate(Rs("PassDate")) Then Rs("PassDate")=Now()
			Rs("Stats")=GroupStats
			Rs("Locked")=Locked
			Rs("ViewFlag")=ViewFlag
			Rs.Update
			Set Rs=Dv_IndivGroup_Conn.Execute("Select GroupID From Dv_GroupUser where GroupID="&GroupID&" And UserID="&AppUserID)
            If Rs.EOF And Rs.BOF Then
			    Dv_IndivGroup_Conn.Execute("Insert Into Dv_GroupUser(GroupID,UserID,UserName,IsLock) Values("&GroupID&","&AppUserID&",'"&AppUserName&"',2)")
            End If
		End If
		Rs.Close:Set Rs=Nothing
		Dv_suc("<b>�޸ĳɹ���</b>")
	End if
	If GroupStats>0 And GroupID>0 And InStr(UserGroup,","&GroupID&",")=0 Then
		If Right(UserGroup,1)="," Then
			UserGroup = UserGroup & GroupID &","
		Else
			UserGroup = UserGroup &","& GroupID &","
		End If
		Dvbbs.Execute("Update Dv_User Set UserGroup='"&UserGroup&"' Where UserID="&AppUserID)
	End If
End Sub

'ɾ��Ȧ��
Sub Del()
	Dim Rs,SQL,GroupID,GroupUserIDStr
	GroupID = Dvbbs.CheckNumeric(Request("groupid"))
	If GroupID>0 Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select * From Dv_Group_Board Where RootID="&GroupID)
		If Not Rs.Eof Then Errmsg=ErrMsg + "<BR><li>��Ȧ������Ŀ���ڣ�����ɾ�����뵽ǰ̨��������Ŀɾ�����ڽ��д˲�����"
	Else
		Errmsg=ErrMsg + "<BR><li>Ȧ��ID������ȷ���Ƿ��ⲿ�ύ��"
	End If
	If Errmsg<>"" Then
		dvbbs_error()
		Exit sub
	End If
	GroupUserIDStr = ""
	Set Rs=Dv_IndivGroup_Conn.Execute("Select UserID From Dv_GroupUser Where GroupID="&GroupID)
	Do While Not Rs.Eof
		If GroupUserIDStr="" Then
			GroupUserIDStr = Rs(0)
		Else
			GroupUserIDStr = GroupUserIDStr&","&Rs(0)
		End If
		Rs.MoveNext
	Loop
	Rs.Close
	If GroupUserIDStr <> "" Then
		Set Rs = Dvbbs.Execute("Select UserID,UserGroup From Dv_User Where UserID In ("&GroupUserIDStr&")")
		Do While Not Rs.Eof
			Dvbbs.Execute("Update Dv_User Set UserGroup='"&Replace(Rs(1),","&GroupID&",",",")&"' Where UserID="&Rs(0))
			Rs.MoveNext
		Loop
		Rs.Close:Set Rs=Nothing
		'Dvbbs.Execute("Update Dv_User Set UserGroup=Replace(UserGroup,',"&GroupID&",',',') Where UserID In ("&GroupUserIDStr&")")
	End If
	Dv_IndivGroup_Conn.Execute("Delete From Dv_GroupUser Where GroupID="&GroupID)
	Dv_IndivGroup_Conn.Execute("Delete From Dv_GroupName Where ID="&GroupID)
	Dv_suc("<b>ɾ���ɹ���</b>")
End Sub

'Ȧ���û��б�
Sub GroupUserList()
	'��ȡȦ������
	Dim Groupid,GroupName
	Dim Rs,SQL,keyword,QuerySty
	GroupID = Dvbbs.CheckNumeric(Request("groupid"))
	keyword = Dvbbs.Checkstr(trim(Request("keyword")))
	GroupName = Dvbbs.Checkstr(trim(Request("GroupName")))
	SQL="":QuerySty=""

	PageSearch = "action=groupuserlist"
	If keyword="" And GroupName="" Then
		If GroupID>0 Then
			PageSearch = PageSearch & "&groupid="&GroupID
			QuerySty = "Where GroupID="&GroupID
			Set Rs=Dv_IndivGroup_Conn.Execute("Select GroupName From Dv_GroupName Where ID="&GroupID)
			If Not Rs.Eof Then GroupName=Rs(0) Else GroupName="δ֪"
			Rs.Close:Set Rs=Nothing
		End If
	Else
		If GroupName<>"" Then
			PageSearch = PageSearch & "&GroupName="&GroupName
			Set Rs=Dv_IndivGroup_Conn.Execute("Select ID From Dv_GroupName Where GroupName='"&GroupName&"'")
			If Not Rs.Eof Then
				QuerySty="Where GroupID="&Rs(0)
			Else
				QuerySty="Where GroupID=0"
			End if
			Rs.Close:Set Rs=Nothing
			If keyword<>"" Then
				PageSearch = PageSearch & "&keyword="&keyword
				QuerySty=QuerySty&" And UserName like '%"&keyword&"%'"
			End if
		Else
			PageSearch = PageSearch & "&keyword="&keyword
			QuerySty="Where UserName like '%"&keyword&"%'"
		End If
	End If
	SQL = "Select ID,GroupID,UserID,UserName,Islock,Intro From Dv_GroupUser "&QuerySty&" Order By ID Desc"
	'Dv_GroupUser
	'islock
	'0=������1=��ˣ�2=����
	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString,i
	Endpage=0:MaxRows=20:CountNum=0
	Page = Dvbbs.CheckNumeric(Request("Page"))
	If Page=0 Then Page=1
%>
<br/>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<th style="text-align:center;" colspan="4"><b>Ȧ�ӳ�Ա����</b></th>
	</tr>
	<form method="post" action="indivGroup.asp?action=groupuserlist" name="searchform">
	<tr>
		<td class=td1>
		&nbsp;��Ա���������ؼ��֣�<input name="keyword" type="text" value="<%=keyword%>" size="15" />
		&nbsp;����Ȧ�����ƣ�<input name="GroupName" type="text" value="<%=GroupName%>" size="15" />
		&nbsp;<input name="Submit" type="submit" class="button" value="����" />

		</td>
	</tr>
	</form>
</table>
<br />
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<form name="theform" method="post" action="indivGroup.asp?action=manage">
	<tr>
		<th style="text-align:center;width:3%">&nbsp;</th>
		<th style="text-align:center;">�û�����</th>
		<th style="text-align:center;">�û���ע</th>
		<th style="text-align:center;">����Ȧ������</th>
		<th style="text-align:center;">�û�״̬</th>
		<th style="text-align:center;">����</th>
	</tr>
<%
	Set Rs = Dvbbs.iCreateObject ("adodb.recordset")
	Rs.Open SQL,Dv_IndivGroup_Conn,1,1
	If Not Rs.eof Then
		CountNum = Rs.RecordCount
		If CountNum Mod MaxRows=0 Then
			Endpage = CountNum \ MaxRows
		Else
			Endpage = CountNum \ MaxRows+1
		End If
		Rs.MoveFirst
		If Page > Endpage Then Page = Endpage
		If Page < 1 Then Page = 1
		If Page >1 Then Rs.Move (Page-1) * MaxRows
		SQL=Rs.GetRows(MaxRows)
		Rs.close:Set Rs = Nothing

		'ID=0,GroupID=1,UserID=2,UserName=3,Islock=4,Intro=5
		For i=0 To Ubound(SQL,2)
%>
	<tr>
		<td class="td1"><input type="checkbox" class="checkbox" name="lid" value="<%=Sql(0,i)%>"></td>
		<td class="td1"><a href="../Dispuser.asp?id=<%=SQL(2,i)%>" target="_blank"><%=SQL(3,i)%></a></td>
		<td class="td1"><%=SQL(5,i)%>&nbsp;</td>
			<%
			If GroupID=0 Then
				Set Rs=Dv_IndivGroup_Conn.Execute("Select GroupName From Dv_GroupName Where ID="&SQL(1,i))
				If Not Rs.Eof Then GroupName=Rs(0) Else GroupName="δ֪"
				Rs.Close:Set Rs=Nothing
			End If
			%>
		<td class="td1"><%=GroupName%></td>
		<td class="td1" align="center">
			<select name="userstats_<%=Sql(0,i)%>">
				<option value="0">���</option>
				<option value="1">����</option>
				<option value="2">����Ա</option>
			</select>
		</td>
		<SCRIPT LANGUAGE="JavaScript">ChkSelected(document.theform.userstats_<%=Sql(0,i)%>,'<%=Sql(4,i)%>');</SCRIPT>
		<td class="td1" align="center">
			<a href="indivGroup.asp?action=useraddedit&groupid=<%=SQL(1,i)%>&groupuserid=<%=SQL(0,i)%>">�޸�</a> |
			<a href="indivGroup.asp?action=manage&act=del&groupid=<%=SQL(1,i)%>&groupuserid=<%=SQL(0,i)%>">ɾ��</a>
		</td>
	</tr>
<%		Next	%>
	<tr>
		<td class="td2" colspan="6">
		��ѡ���û���<input type="checkbox" class="checkbox" name="chkall" value="on" onclick="CheckAll(this.form)">ȫѡ
		<input type="submit" class="button" name="act" value="����ɾ��"  onclick="{if(confirm('��ȷ��Ҫɾ����ѡ��ȫ��Ȧ�ӳ�Ա��?')){this.document.theform.submit();return true;}return false;}">
		<input type="submit" class="button" name="act" value="��������" onclick="{if(confirm('��ȷ��Ҫ������ѡ�ļ�¼��?')){this.document.theform.submit();return true;}return false;}">
		</td>
	</tr>
<%
	Else
		Rs.close:Set Rs = Nothing
		Response.write "<tr><td class=""td1"" colspan=""6"" align=""center"">û���������κ��û����ݣ�</td></tr>"
	End If
%>
	<tr><td class="td1" colspan="6"><SCRIPT language="javascript">PageList('<%=Page%>',10,'<%=MaxRows%>','<%=CountNum%>','<%=PageSearch%>',1);</SCRIPT>&nbsp;</td></tr>
	</form>
</table>
<%
End Sub

Sub UserManage()
	'Ȧ���û�����
	Dim Rs,SQL,i,GroupID,GroupUserID,UserStats
	'Response.write Request.ServerVariables("QUERY_STRING")
	'Response.end
	If LCase(Request("act"))="del" Then
		GroupID = Dvbbs.CheckNumeric(Request("GroupID"))
		GroupUserID = Dvbbs.CheckNumeric(Request("GroupUserID"))
		Set Rs=Dv_IndivGroup_Conn.Execute("Select GroupID,UserID From Dv_GroupUser Where ID="&GroupUserID)
		If Not Rs.Eof Then Updategroupname Rs(1),GroupID,0
		Dv_IndivGroup_Conn.Execute("Delete From Dv_GroupUser Where ID="&GroupUserID)
	Else
		If Request.Form("lid")="" Then
			Errmsg=ErrMsg + "��ָ������û���"
			dvbbs_error()
			Exit Sub
		Else
			GroupUserID=Replace(Request.Form("lid"),"'","")
			GroupUserID=Replace(GroupUserID,";","")
			GroupUserID=Replace(GroupUserID,"--","")
			GroupUserID=Replace(GroupUserID,")","")
		End If

		If Request("act")="����ɾ��" Then
			SQL = "Select GroupID,UserID From Dv_GroupUser Where ID In ("&GroupUserID&")"
			Set Rs = Dv_IndivGroup_Conn.Execute(SQL)
			If Not Rs.Eof Then
				SQL = Rs.GetRows(-1)
				For i=0 to Ubound(Sql,2)
					Updategroupname SQL(1,i),SQL(0,i),0
				Next
				Dv_IndivGroup_Conn.Execute("Delete From Dv_GroupUser Where ID In ("&GroupUserID&")")
			End If
			Rs.Close:Set Rs=nothing
		ElseIf Request("act")="��������" Then
			SQL = "Select id,GroupID,UserID,UserName,IsLock From Dv_GroupUser Where ID In ("&GroupUserID&")"
			Set Rs=Dvbbs.iCreateObject("Adodb.RecordSet")
			Rs.Open SQL,Dv_IndivGroup_Conn,1,3
			Do While Not Rs.Eof
				UserStats = Dvbbs.CheckStr(Request.Form("userstats_"&Rs(0)))
				If UserStats<>Rs(4) Then
					If UserStats=0 Then Updategroupname Rs(2),Rs(1),0
					If Rs(4)=0 Then Updategroupname Rs(2),0,Rs(1)
					Rs(4) = Dvbbs.CheckStr(Request.Form("userstats_"&Rs(0)))
					Rs.Update
				End If
				Rs.Movenext
			Loop
			Rs.Close:Set Rs=Nothing
		End If
	End If
	Dv_suc(Request("act")&"<b>�ɹ�ִ��!</b>")
End Sub

'��ID���߼�1 ��ID�ƽ���1
Sub Updategroupname(UserID,old_gid,new_id)
	Dim Rs,SQL,UserGroup
	If old_gid>0 Then Dv_IndivGroup_Conn.Execute("Update Dv_GroupName Set UserNum=UserNum-1 Where ID="&old_gid&"")
	If new_id>0 Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select UserNum,limituser From Dv_GroupName Where ID="&new_id&"")
		If Not Rs.Eof Then
			SQL = "UserNum="&Rs(0)+1
			If Rs(1)=Rs(0)+1 Then SQL = SQL & ",IsLock=2"
			Dv_IndivGroup_Conn.Execute("Update Dv_GroupName Set "&SQL&" Where ID="&new_id&"")
		End If
		Rs.Close:Set Rs = Nothing
	End If
	'Ȧ�ӳ�Ա���������롢ɾ��
	Set Rs=Dv_IndivGroup_Conn.Execute("Select UserGroup From Dv_user Where UserID="&UserID)
	If Not Rs.Eof Then UserGroup = Rs(0)
	If InStr(UserGroup,",")=0 Or IsNull(InStr(UserGroup,",")) Then UserGroup=","
	If old_gid>0 and new_id>0 Then
		UserGroup = Replace(UserGroup,","&old_gid&",",",")
		UserGroup = UserGroup & new_id&","
		Dv_IndivGroup_Conn.Execute("Update Dv_user Set UserGroup = '"&UserGroup&"' Where UserID="&UserID)
	Else
		If new_id>0 Then
			UserGroup = UserGroup & new_id&","
			Dv_IndivGroup_Conn.Execute("Update Dv_user Set UserGroup = '"&UserGroup&"' Where UserID="&UserID)
		Else
			UserGroup = Replace(UserGroup,","&old_gid&",",",")
			Dv_IndivGroup_Conn.Execute("Update Dv_user Set UserGroup = '"&UserGroup&"' Where UserID="&UserID)
		End If
	End If
End Sub

'Ȧ��״̬
Function GroupStatsStr(Sid)
	Select Case Sid
		Case 1 : GroupStatsStr = "����"
		Case 2 : GroupStatsStr = "����"
		Case 3 : GroupStatsStr = "�ر�"
		Case 0 : GroupStatsStr = "���"
		Case Else : GroupStatsStr = "δ֪"
	End Select
End Function

'Ȧ���û�״̬
Function GroupUserStats(Sid)
	Select Case Sid
		Case 1 : GroupUserStats = "����"
		Case 2 : GroupUserStats = "����Ա"
		Case 0 : GroupUserStats = "���"
		Case Else : GroupUserStats = "δ֪"
	End Select
End Function


Sub LoadGroupClass()
	Dim SqlStr,Rs
	SqlStr="Select * From Dv_Group_Class"
	Set Rs=Dvbbs.Execute(SqlStr)
	Set Application(Dvbbs.CacheName & "_Dv_Group_Class")=Dvbbs.RecordsetToxml(Rs,"list","groupclass")
	Rs.Close
End Sub
%>