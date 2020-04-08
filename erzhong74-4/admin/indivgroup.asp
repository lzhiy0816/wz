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
		Case "更新用户统计" : upusernum()
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
		<th width="100%" style="text-align:center;" colspan=2>论坛圈子管理</th>
	</tr>
	<tr>
		<td height="23" colspan="2" class=td1>请在插件管理的菜单管理中添加一级菜单“个性圈子”。</td>
	</tr>
	<tr>
		<td height="23" colspan="2" class=td1><b>用户圈子管理</b>：您可以添加修改或者删除论坛圈子。</td>
	</tr>
	<tr>
		<td height="23" colspan="2" class=td1>
			<button onclick="window.location='indivGroup.asp?action=addedit&orders=0'" class="button">添加圈子</button>&nbsp;
			<button onclick="window.location='indivGroup.asp'" class="button">圈子管理</button>&nbsp;
			<button onclick="window.location='indivGroup.asp?action=groupuserlist'" class="button">圈子用户管理</button>&nbsp;
			<button onclick="window.location='indivGroup.asp?action=groupclasslist'" class="button">圈子分类管理</button>&nbsp;
		</td>
	</tr>
</table>
<br/>
<%
End Sub

'圈子列表
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
			Response.write "<script language='javascript'>alert('请选择要修改的项');</script>"
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
			Response.write "<script language='javascript'>alert('修改成功！');</script>"
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
		<th width="100%" style="text-align:center;" colspan=2><b>添加/修改圈子</b></th>
	</tr>
	<tr>
		<td width="20%" class=td1><b>圈子名称</b></td>
		<td width="80%" class=td1>
		<input type="text" name="Groupname" size="35" value="<%=Groupname%>">
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>管理员名称</b></td>
		<td width="80%" class=td1>
		<input type="text" name="AppUserName" size="35" value="<%=AppUserName%>">
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>圈子说明</b></td>
		<td width="80%" class=td1>
		<textarea rows="4" cols="35" name="GroupInfo"><%=GroupInfo%></textarea>
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>成员人数上限</b></td>
		<td width="80%" class=td1>
		<input type="text" name="LimitUser" size="35" value="<%=LimitUser%>">
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>圈子状态</b></td>
		<td width="80%" class=td1>
		<%If GroupStats=0 Then Response.write "<input type=""radio"" class=""radio"" name=""GroupStats"" value=""0""/>审核"%>
		<input type="radio" class="radio" name="GroupStats" value="1"/>正常
		<input type="radio" class="radio" name="GroupStats" value="2"/>锁定
		<input type="radio" class="radio" name="GroupStats" value="3"/>关闭
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>成员加入</b></td>
		<td width="80%" class=td1>
		<input type="radio" class="radio" name="locked" value="0"/>自由加入
		<input type="radio" class="radio" name="locked" value="1"/>需要审核
		</td>
	</tr>
	<tr>
		<td width="20%" class=td1><b>浏览设置</b></td>
		<td width="80%" class=td1>
		<input type="radio" class="radio" name="viewflag" value="0"/>公开
		<input type="radio" class="radio" name="viewflag" value="1"/>不公开
		</td>
	</tr>
	<tr><td colspan="2">（公开：任何人都可以浏览；不公开：只有成员能浏览）</td></tr>
	<tr><td height="23" colspan="2" class=td1><input type="submit" class="button" name="Submit" value="提 交"></tr>
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
		&nbsp;个性圈子搜索：<input name="keyword" type="text" value="<%=keyword%>" size="15" /> <input name="Submit" type="submit" class="button" value="搜索" />
		<select name="orders" onchange='javascript:submit()'>
			<option value="0" selected>所有圈子</option>
			<option value="1">待审核圈子</option>
			<option value="2">最新圈子</option>
			<option value="3">发贴总数Top20</option>
			<option value="4">成员总数Top20</option>
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
		<th style="text-align:center;">圈子名称</th>
		<th style="text-align:center;">圈子申请人名称</th>
		<th style="text-align:center;">用户总数</th>
		<th style="text-align:center;">帖子总数</th>
		<th style="text-align:center;">主题总数</th>
		<th style="text-align:center;">成员人数上限</th>
		<th style="text-align:center;">申请时间</th>
		<th style="text-align:center;">状态</th>
		<th style="text-align:center;">操作</th>
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
			<option value="1">正常</option>
			<option value="2">锁定</option>
			<option value="3">关闭</option>
			<option value="0">审核</option>
		</select>
		<script language="javascript">ChkSelected(document.batchform.GroupStats_<%=SQL(0,i)%>,'<%=SQL(6,i)%>');</script>
		<%=GroupStatsStr(SQL(6,i))%>
		</td>
		<td height="23" class="td2" align="right">
			<%if SQL(6,i)=0 Then Response.write "<a href=""indivGroup.asp?groupid="&SQL(0,i)&"&action=pass&"&QueryString&""">通过</a> | "%>
			<a href="indivGroup.asp?groupid=<%=SQL(0,i)%>&action=addedit&<%=QueryString%>">修改</a> |
			<a href="indivGroup.asp?groupid=<%=SQL(0,i)%>&action=del&<%=QueryString%>">删除</a> |
			<a href="indivGroup.asp?action=GroupUserList&groupid=<%=SQL(0,i)%>">查看用户</a>
		</td>
	</tr>
<%
		Next
	Else
		If orders=1 Then
			Response.write "<tr><td class=""td1"" colspan=""10"">没有待审核的圈子</td></tr>"
		Else
			Response.write "<tr><td class=""td1"" colspan=""10"">没找到满足搜索条件的圈子</td></tr>"
		End If
		Rs.Close:Set Rs=Nothing
	End If
%>
	<tr>
		<td colspan="10">
			<div style="margin-left:35px;">
			<input type="submit" name="Submit" value="批量修改" onclick="return clicksubmit(this.form);" class="button"/>
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);" class="chkbox" />&nbsp;全选/取消
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
	if(confirm('您确定执行的操作吗?')){
		this.document.batchform.submit();
		return true;
	}else{
		return false;
	}
}
</script>
<%
End Sub

Sub GroupClassList() Rem 圈子分类
	Dim Sql,SqlStr,Rs
	Set Rs = Dvbbs.Execute("Select ID,ClassName,GroupCount,Orders From Dv_Group_Class Order By Orders,Id")
	%>
	<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
		<form method="post" action="indivGroup.asp?action=UpdateGroupClass" name="searchform">
		<tr>
			<td width="5%" style="text-align:center;">选择</td>
			<td width="5%" style="text-align:center;">ID</td>
			<td style="text-align:center;">圈子分类名称</td>
			<td width="10%" style="text-align:center;">排序 ↓</td>
			<td width="10%" style="text-align:center;">圈子数量</td>
			<td width="10%" style="text-align:center;">分类操作</td>
		</tr>
		<%Do While Not Rs.EOF%>
		<tr>
			<td class="td1" style="text-align:center;"><input type="hidden" name="Id" value="<%=Rs(0)%>"/>&nbsp;</td>
			<td class="td1" style="text-align:center;"><%=Rs(0)%></td>
			<td class="td1"><input type="text" size="35" name="ClassName" value="<%=Rs(1)%>" /></td>
			<td class="td1"><input type="text" size="4" name="Orders" maxlength="3" value="<%=Rs(3)%>" /></td>
			<td class="td1" style="text-align:center;"><a href="../IndivGroup_List.asp?ClassId=<%=Rs(0)%>" target="_blank"><%=Rs(2)%></a></td>
			<td width="10%" style="text-align:center;"><a href="indivGroup.asp?action=DeleteGroupClass&Id=<%=Rs(0)%>" onclick="{if(confirm('你确定要删除该圈子分类 [<%=Rs(1)%>] 吗？，删除操作不可恢复！')){return true;}return false;}">删除分类</a></td>
		</tr>
		<%
		Rs.MoveNext
		Loop
		%>
		<tr>
			<td class="td1" style="text-align:center;"><input type="checkbox" name="addClass" value="1"/></td>
			<td class="td1" style="text-align:center;"><font color="#FF0000">新增</font></td>
			<td class="td1"><input type="text" size="35" name="addClassName" /></td>
			<td class="td1"><input type="text" size="4" name="addOrders" maxlength="3" value="0" /></td>
			<td class="td1" style="text-align:center;">&nbsp;</td>
			<td width="10%" style="text-align:center;">&nbsp;</td>
		</tr>
		<tr>
		<td colspan="6" style="text-align:center;">
			<input type="checkbox" name="chkall" value="on" onclick="CheckAll(this.form);" class="chkbox" />&nbsp;全选/取消&nbsp;&nbsp;
			<input type="submit" name="Submit" value="批量修改" onclick="" class="button"/>
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
			Str = "<font color=""#FF0000"">Id:"&Id&" 圈子名称不能为空，排序必须是数字！</font>"
			Exit For
		End If
	Next

	If CInt(Request.Form("addClass"))=1 Then
		ClassName = Trim(Replace(Request.Form("addClassName"),"'",""))
		Orders = Cint(Request.Form("addOrders"))
		Set Rs = Dvbbs.Execute("Select ID,ClassName From Dv_Group_Class Where ClassName='"&ClassName&"'")
		If Not Rs.Eof Then
			Str = "<br /><font color=""#FF0000"">Error: 已经存在新增圈子分类的名称,请重新填写！</font>"
		Else
			If ClassName<>"" Then
				Dvbbs.Execute("Insert Into [Dv_Group_Class] (ClassName,GroupCount,Orders) Values('"&ClassName&"',0,"&Orders&")")
				Str = "<br /><font color=""#FF0000"">新增圈子分类 <b>"&ClassName&"</b> 成功！</font>"
			Else
				Str = "<br /><font color=""#FF0000"">Error: 新增圈子分类,请填写新圈子分类名称！</font>"
			End If
		End If
	End If
	LoadGroupClass()
	Dv_suc("批量圈子分类数据成功！"& Str)
End Sub

Sub DeleteGroupClass()
	Dim Id
	Id=CLng(Request("Id"))
	Dvbbs.Execute("Update Dv_GroupName Set ClassId=0 Where ClassId="&Id)
	Dvbbs.Execute("Delete From Dv_Group_Class Where Id="&Id)
	Dv_suc("删除圈子分类成功！原下属圈子已经取消分类。")
End Sub

'保存圈子信息
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
		If GroupName = "" Then Errmsg=ErrMsg + "<BR><li>名称不能为空。"
		If AppUserName = "" Then Errmsg=ErrMsg + "<BR><li>圈子管理员不能为空。"
		QueryStr = ""
	Else
		QueryStr = "And ID<>"&GroupID
	End If

	If GroupName<>"" Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select * From Dv_GroupName Where GroupName='"&GroupName&"' "&QueryStr&" Order By ID Desc")
		If Not Rs.Eof Then Errmsg=ErrMsg + "<BR><li>该圈子名称已经被申请，请填写其他圈子名称。"
		Rs.Close:Set Rs=Nothing
	End If
	If AppUserName<>"" Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select * From Dv_GroupName Where AppUserName='"&AppUserName&"' "&QueryStr&" Order By ID Desc")
		If Not Rs.Eof Then Errmsg=ErrMsg + "<BR><li>这个用户已经申请有圈子了，一个用户只能申请一个圈子。"
		Rs.Close:Set Rs=Nothing

		Set Rs=Dv_IndivGroup_Conn.Execute("Select UserID,UserGroupID,UserGroup From Dv_user Where UserName='"&AppUserName&"'")
		If Not Rs.Eof Then
			If Rs(1)<>5 Then
				AppUserID = Rs(0)
				UserGroup = Dvbbs.CheckStr(Rs(2))
			Else
				Errmsg=ErrMsg + "<BR><li>用户 "&AppUserName&" 是等待验证的(COPPA)会员，不能作为圈子管理员。"
			End If
		Else
			Errmsg=ErrMsg + "<BR><li>用户 "&AppUserName&" 还没有注册，不能作为圈子管理员。"
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
		Dv_suc("<b>添加成功！</b>")
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
		Dv_suc("<b>修改成功！</b>")
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

'删除圈子
Sub Del()
	Dim Rs,SQL,GroupID,GroupUserIDStr
	GroupID = Dvbbs.CheckNumeric(Request("groupid"))
	If GroupID>0 Then
		Set Rs=Dv_IndivGroup_Conn.Execute("Select * From Dv_Group_Board Where RootID="&GroupID)
		If Not Rs.Eof Then Errmsg=ErrMsg + "<BR><li>该圈子有栏目存在，不能删除，请到前台将所有栏目删除后在进行此操作。"
	Else
		Errmsg=ErrMsg + "<BR><li>圈子ID错误，请确认是否外部提交。"
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
	Dv_suc("<b>删除成功！</b>")
End Sub

'圈子用户列表
Sub GroupUserList()
	'获取圈子属性
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
			If Not Rs.Eof Then GroupName=Rs(0) Else GroupName="未知"
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
	'0=正常，1=审核，2=管理
	Dim Page,MaxRows,Endpage,CountNum,PageSearch,SqlString,i
	Endpage=0:MaxRows=20:CountNum=0
	Page = Dvbbs.CheckNumeric(Request("Page"))
	If Page=0 Then Page=1
%>
<br/>
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<tr>
		<th style="text-align:center;" colspan="4"><b>圈子成员搜索</b></th>
	</tr>
	<form method="post" action="indivGroup.asp?action=groupuserlist" name="searchform">
	<tr>
		<td class=td1>
		&nbsp;成员名称搜索关键字：<input name="keyword" type="text" value="<%=keyword%>" size="15" />
		&nbsp;所属圈子名称：<input name="GroupName" type="text" value="<%=GroupName%>" size="15" />
		&nbsp;<input name="Submit" type="submit" class="button" value="搜索" />

		</td>
	</tr>
	</form>
</table>
<br />
<table width="100%" border="0" cellspacing="1" cellpadding="3" align="center">
	<form name="theform" method="post" action="indivGroup.asp?action=manage">
	<tr>
		<th style="text-align:center;width:3%">&nbsp;</th>
		<th style="text-align:center;">用户名称</th>
		<th style="text-align:center;">用户备注</th>
		<th style="text-align:center;">所属圈子名称</th>
		<th style="text-align:center;">用户状态</th>
		<th style="text-align:center;">操作</th>
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
				If Not Rs.Eof Then GroupName=Rs(0) Else GroupName="未知"
				Rs.Close:Set Rs=Nothing
			End If
			%>
		<td class="td1"><%=GroupName%></td>
		<td class="td1" align="center">
			<select name="userstats_<%=Sql(0,i)%>">
				<option value="0">审核</option>
				<option value="1">正常</option>
				<option value="2">管理员</option>
			</select>
		</td>
		<SCRIPT LANGUAGE="JavaScript">ChkSelected(document.theform.userstats_<%=Sql(0,i)%>,'<%=Sql(4,i)%>');</SCRIPT>
		<td class="td1" align="center">
			<a href="indivGroup.asp?action=useraddedit&groupid=<%=SQL(1,i)%>&groupuserid=<%=SQL(0,i)%>">修改</a> |
			<a href="indivGroup.asp?action=manage&act=del&groupid=<%=SQL(1,i)%>&groupuserid=<%=SQL(0,i)%>">删除</a>
		</td>
	</tr>
<%		Next	%>
	<tr>
		<td class="td2" colspan="6">
		请选择用户，<input type="checkbox" class="checkbox" name="chkall" value="on" onclick="CheckAll(this.form)">全选
		<input type="submit" class="button" name="act" value="批量删除"  onclick="{if(confirm('您确定要删除所选的全部圈子成员吗?')){this.document.theform.submit();return true;}return false;}">
		<input type="submit" class="button" name="act" value="批量更改" onclick="{if(confirm('您确定要更改所选的记录吗?')){this.document.theform.submit();return true;}return false;}">
		</td>
	</tr>
<%
	Else
		Rs.close:Set Rs = Nothing
		Response.write "<tr><td class=""td1"" colspan=""6"" align=""center"">没有搜索到任何用户数据！</td></tr>"
	End If
%>
	<tr><td class="td1" colspan="6"><SCRIPT language="javascript">PageList('<%=Page%>',10,'<%=MaxRows%>','<%=CountNum%>','<%=PageSearch%>',1);</SCRIPT>&nbsp;</td></tr>
	</form>
</table>
<%
End Sub

Sub UserManage()
	'圈子用户管理
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
			Errmsg=ErrMsg + "请指定相关用户。"
			dvbbs_error()
			Exit Sub
		Else
			GroupUserID=Replace(Request.Form("lid"),"'","")
			GroupUserID=Replace(GroupUserID,";","")
			GroupUserID=Replace(GroupUserID,"--","")
			GroupUserID=Replace(GroupUserID,")","")
		End If

		If Request("act")="批量删除" Then
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
		ElseIf Request("act")="批量更改" Then
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
	Dv_suc(Request("act")&"<b>成功执行!</b>")
End Sub

'旧ID移走减1 新ID移进加1
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
	'圈子成员更换、加入、删除
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

'圈子状态
Function GroupStatsStr(Sid)
	Select Case Sid
		Case 1 : GroupStatsStr = "正常"
		Case 2 : GroupStatsStr = "锁定"
		Case 3 : GroupStatsStr = "关闭"
		Case 0 : GroupStatsStr = "审核"
		Case Else : GroupStatsStr = "未知"
	End Select
End Function

'圈子用户状态
Function GroupUserStats(Sid)
	Select Case Sid
		Case 1 : GroupUserStats = "正常"
		Case 2 : GroupUserStats = "管理员"
		Case 0 : GroupUserStats = "审核"
		Case Else : GroupUserStats = "未知"
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