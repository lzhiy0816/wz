<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dim admin_flag,CssList,StyleConn
admin_flag=",21,"
If not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
	Head()
	Errmsg=ErrMsg + "<br><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。</li><br><li>您没有管理本页面的权限。<.li>"
	dvbbs_error()
	Call Footer()
	Response.End
End If
 Select Case Request("action")
 	Case ""
 		Main()
 	Case "doout"
 		doout()
 	Case "load"
 		 Load()
 	 	Case "doload"
 		 doLoad()
 	Case "doupdate"
 		 doupdate()
 End Select
Sub doupdate()
	Dim Rs,node,inid,cssid,RsSkin,i,cssdom,toid
	If Request.form("submit")="" Then
		Head()
		Readme()
		Setup1
	Else
		If Request.form("inid")="" Then
			Head()
			Errmsg=ErrMsg + "<BR><li>必须选择源模板."
			dvbbs_error()
			Call Footer()	
			Exit Sub
		End If
		If Request.form("toid")="" Then
			Head()
			Errmsg=ErrMsg + "<BR><li>必须选择目标模板."
			dvbbs_error()
			Call Footer()	
			Exit Sub
		End If
		If SkinConnection(Request.Form("skinmdb")) Then
			inid=Request.form("inid")
			toid=Request.form("toid")
			Set Rs=StyleConn.Execute("select * From Dv_style Where ID="&CLng(inid))
			Set RsSkin=Server.CreateObject("adodb.recordset")
			RsSkin.open "Select * From Dv_Style Where Id="& CLng(toid),Conn,1,3
			For i=2 to 19
			If  Request.form(RsSkin(i).Name)="1" Then
				RsSkin(i)=Rs(RsSkin(i).Name)
			End If
			Next
			RsSkin.update
			RsSkin.Close
			If Request.form("Forum_CSS")="1" Then
				Set CssList=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				Set cssdom=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				CssList.LoadXML Rs("Forum_CSS")
				Set Rs=Dvbbs.Execute("select Forum_Css From Dv_Setup")
				cssdom.LoadXMl Rs(0)
				SetTid toid
				Rem 删除该模板的CSS样式
				For Each Node in cssdom.documentElement.selectNodes("css[tid='"&toid&"']")
					cssdom.documentElement.removeChild(node)
				Next
				For Each Node in CssList.documentElement.selectNodes("css")
					cssdom.documentElement.appendChild(node)
				Next
					i=1
					For Each Node in cssdom.documentElement.selectNodes("css/@filename")
						Node.text="aspsky_"&i
						i=i+1
					Next
					i=1
					For Each Node in cssdom.documentElement.selectNodes("css/@id")
						Node.text=i
						i=i+1
					Next
					Dvbbs.Execute("Update Dv_Setup Set Forum_Css='"&Dvbbs.Checkstr(cssdom.xml)&"'")
			End If
			Dvbbs.loadSetup()
			Dvbbs.Loadstyle()
			createsccfile()
			Head()
			Dv_suc("模板覆盖更新完成！")
			Call Footer()	
		Else
			Head()
			Errmsg=ErrMsg + "<BR><li>目标数据库"& Request.Form("skinmdb")&"不存在或有错误.</li>"
			dvbbs_error()
			Call Footer()	
		End If
	End If
End Sub
Sub Load()
	If Request("setup")="" or Request.form("submit")="" Then
		Head()
		Readme()
		Setup1
	ElseIf Request("setup")="1" Then
		Head()
		Setup2
	End If
	Footer()
End Sub
Sub doload()
		Dim Rs,node,skid,cssid,RsSkin,i,cssdom
		If Request.form("skid")="" Then
				Head()
			Errmsg=ErrMsg + "<BR><li>必须选择要导入的模板."
			dvbbs_error()
			Call Footer()	
			Exit Sub
		End If
	 If SkinConnection(Request.Form("skinmdb")) Then
		Set Rs=Dvbbs.Execute("select Forum_Css From Dv_Setup")
		Set CssList=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		Set cssdom=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		cssdom.LoadXMl Rs(0)
		Set Rs=Nothing
		For Each skid in Request.form("skid") 
			Set Rs=StyleConn.Execute("select * From Dv_style Where ID="&CLng(skid))
			Set cssid=Request.Form("cssid_"&skid)
			Set RsSkin=Server.CreateObject("adodb.recordset")
			RsSkin.open "Select * From Dv_Style Where Id=0",Conn,1,3
			RsSkin.AddNew
			RsSkin("StyleName")=Request.Form("StyleName"&skid)
			RsSkin("Main_Style")=Rs("Main_Style")
			RsSkin("Style_Pic")=Rs("Style_Pic")
			RsSkin("page_index")=Rs("page_index")
			RsSkin("page_dispbbs")=Rs("page_dispbbs")
			RsSkin("page_showerr")=Rs("page_showerr")
			RsSkin("page_login")=Rs("page_login")
			RsSkin("page_online")=Rs("page_online")
			RsSkin("page_usermanager")=Rs("page_usermanager")
			RsSkin("page_fmanage")=Rs("page_fmanage")
			RsSkin("page_boardstat")=Rs("page_boardstat")
			RsSkin("page_paper_even_toplist")=Rs("page_paper_even_toplist")
			RsSkin("page_query")=Rs("page_query")
			RsSkin("page_show")=Rs("page_show")
			RsSkin("page_dispuser")=Rs("page_dispuser")
			RsSkin("page_help_permission")=Rs("page_help_permission")
			RsSkin("page_postjob")=Rs("page_postjob")
			RsSkin("page_post")=Rs("page_post")
			RsSkin("page_boardhelp")=Rs("page_boardhelp")
			CssList.LoadXML Rs("Forum_CSS")
			Set CssList=OutCSSDom(cssid)
			RsSkin.Update
			RsSkin.MoveLast
			SetTid(RsSkin("id"))
			RsSkin.Close
			For Each Node in CssList.documentElement.selectNodes("css")
				cssdom.documentElement.appendChild(node)
			Next
		Next
		i=1
		For Each Node in cssdom.documentElement.selectNodes("css/@filename")
			Node.text="aspsky_"&i
			i=i+1
		Next
		i=1
		For Each Node in cssdom.documentElement.selectNodes("css/@id")
			Node.text=i
			i=i+1
		Next
		Dvbbs.Execute("Update Dv_Setup Set Forum_Css='"&Dvbbs.Checkstr(cssdom.xml)&"'")
		Dvbbs.loadSetup()
		Dvbbs.Loadstyle()
		createsccfile()
		Head()
		Dv_suc("模板成功导入")
		Call Footer()	
	 Else
	 		Head()
			Errmsg=ErrMsg + "<BR><li>目标数据库"& Request.Form("skinmdb")&"不存在或有错误.</li>"
			dvbbs_error()
			Call Footer()	
	 End If
End Sub
Sub createsccfile()
	On error resume Next
	Dim Fso,filename,Forum_CSS
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	If  Err Then
		err.Clear
		Errmsg=ErrMsg + "<br /><li>您的服务器不支持写文件,CSS文件写入失败,请手工操作或把生成文件的内容清空!</li>"
		Dvbbs_error()
		Exit Sub
	End If
	For Each filename In Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css/@filename")
		If  filename.text<>"" Then
			If InStr(filename.text,".")=0 Then
				Dvbbs.SkinID=filename.selectSingleNode("../tid").text
				Dvbbs.LoadTemplates("")
				Forum_CSS=filename.selectSingleNode("../cssdata").text
				Forum_CSS=Replace(Forum_CSS,"{$width}",Dvbbs.mainsetting(0))
				Forum_CSS=Replace(Forum_CSS,"{$PicUrl}",filename.selectSingleNode("../@picurl").text)
				Fso.CreateTextFile(server.MapPath("../skins/"& filename.text &".css")).WriteLine(Forum_CSS)
				If  Err Then
					err.Clear
					Errmsg=ErrMsg + "<br /><li>您的服务器不支持写文件,CSS文件写入失败,请手工操作或把生成文件的内容清空!</li>"
					Dvbbs_error()
					Exit Sub
				End If
			End If
		End If
	Next
	Set FSO=Nothing  
End Sub
Sub SetTid(id)
	Dim Node
	For Each Node in CssList.documentElement.selectNodes("css/tid")
		node.text=id	
	Next
End Sub
Sub Setup2()
	If SkinConnection(Request.Form("skinmdb")) Then
			Dim Rs,node
	Set CssList=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	
	Set Rs=StyleConn.Execute("select ID,StyleName ,Forum_CSS,DateAndTime,readme From Dv_style")
	%>
	<form method="post" action="Loadskin.asp?action=doload">
<table border="0" cellspacing="1" cellpadding="4" align="center" width="95%" class="tableBorder">
	<tr><th colspan="5">模板和CSS的导入[追加模式]</th>
	</tr>
		<tr>
		<td class="forumrow" align="left" colspan="5" style="padding:10px;">您现在要执行的是模板和CSS的导入,把需要导入的勾上再提交</td>
			</tr>
	<tr>
	
	<td width="40"  align="center" class="bodytitle">选择</td>
	<td width="100"  align="center" class="bodytitle">日期</td>
	<td width="100"  align="center" class="bodytitle">模版名称</td>
	<td width="150" class="bodytitle">附加说明</td>
	<td width="302"  class="bodytitle">适用的CSS样式</td>
	</tr>
	<%
	do while not Rs.eof
	CssList.LoadXMl Rs("Forum_CSS")
	%>
	<tr>
	<td width="40" class="forumrow" align="center"><input type="checkbox" name="skid" value="<%=Rs("id")%>"></td>
	<td width="100" class="forumrow" align="center"><%=rs("DateAndTime")%></td>
	<td width="100" class="forumrow" align="center"><Input tyle="text" Name="StyleName<%=Rs("id")%>" value="<%=Rs("StyleName")%>"></td>
	<td width="150" class="forumrow"><div align="left">
	<%=rs("readme")%>
	</div></td>
	<td width="402" class="forumrow">
	<%
	For Each Node in CssList.documentElement.selectNodes("css")
	%>
	<input type="checkbox" name="cssid_<%=Rs("id")%>" value="<%=node.selectSingleNode("@id").text%>">
	<%=node.selectSingleNode("@type").text%>
	<%
	Next
	%>
	</td>
	</tr>
	<%	Rs.movenext
		loop
		Rs.close:Set Rs=Nothing
		Set Rs=StyleConn.Execute("select * From Dv_style")
		Dim Rs1
		Set Rs1=Dvbbs.Execute("select ID,StyleName  From Dv_style")
	%>
	<tr>
	<td colspan="5" align=center class="forumRowHighlight">
	数据库：<input type="text" name="skinmdb" size="30" value="<%=Request.Form("skinmdb")%>" readonly>
	<input type="submit" name="submit" value="导 入" > 
	</td>
	</tr>
	</table>
	</form>
	<br>
	<form method="post" action="Loadskin.asp?action=doupdate">
	<table border="0" cellspacing="1" cellpadding="4" align="center" width="95%" class="tableBorder">
	<tr><th colspan="2">模板和CSS的导入[覆盖模式]</th>
	</tr>
	<tr>
		<td class="forumrow" align="left" colspan="2" style="padding:10px;">请在源模板中选择一份,再在目标(论坛)的模板中选择一份.然后提交,覆盖操作每次只可以处理一份模板.</td>
			</tr>
	<tr>
	<td width="50%"  align="center" class="bodytitle">选择源模板</td>
	<td width="50%"  align="center" class="bodytitle">选择被覆盖的目标模板</td>
	</tr>
	<tr>
	<td width="50%" class="forumrow" align="left" style="padding:5px;">
	<%
	i =0
	do while not Rs.eof%>
	<div><input type="radio" name="inid" value="<%=Rs("ID")%>" <% If i= 0 Then Response.Write " checked"%>><b><%=Rs("StyleName")%></b><ul style="margin-top : 0px;margin-bottom : 0px;"><li><b>日期：</b><%=Rs("DateAndTime")%></li><li><b>说明：</b><%=Rs("readme")%></li></ul></div>
	<%
	i=1
	Rs.movenext
		loop
	%>
	</td>
	<td width="50%" class="forumrow" align="left" style="padding:5px;">
	<%
		i =0
	do while not Rs1.eof%>
	<div><input type="radio" name="toid" value="<%=Rs1(0)%>" <% If i= 0 Then Response.Write " checked"%>><b><%=Rs1(1)%></b></div>
	<%
		i=1
	Rs1.movenext
		loop
	%>
	</td>
	</tr>
	<tr>
	<td width="50%"  align="center" class="bodytitle">数据项目</td>
	<td width="50%"  align="center" class="bodytitle">操作选择</td>
	</tr>
	<% For i= 2 to 20%>
	<tr>
	<td width="50%" class="forumrow" align="center"><%=Rs(i).name%></td>
	<td width="50%" class="forumrow" align="center">
	<input type="checkbox" name="<%=Rs(i).name%>" value="1" checked> 覆盖
	</td>
	</tr>
	<%Next%>
	<tr>
	<td colspan="2" align="center" class="forumRowHighlight">
	数据库：<input type="text" name="skinmdb" size="30" value="<%=Request.Form("skinmdb")%>" readonly>
	<input type="submit" name="submit" value="覆 盖" > 
	</td>
	</tr>
	</table>
	</form>
	<%
	Else
			Errmsg=ErrMsg + "<BR><li>数据库"& Request.Form("skinmdb")&"不存在或有错误.</li>"
			dvbbs_error()
	End If
End Sub
Sub Main
	Head()
	Readme()
	Skinlist()
	Footer()
End Sub
Sub Setup1()%>
		<form method="post" action="Loadskin.asp?action=load&setup=1">
<table border="0" cellspacing="1" cellpadding="4" align="center" width="95%" class="tableBorder">
	<tr><th colspan="4">模板和CSS的入</th>
	</tr>
		<tr>
		<td class="forumrow" align="left" colspan="4" style="padding:10px;">您现在要执行的是模板和CSS的导入,请设置好要导入的源数据库名称,然后提交.</td>
			</tr>
	<tr>
	<td colspan="4" align=center class="forumRowHighlight">
	数据库：<input type="text" name="skinmdb" size="30" value="Dv_skin.mdb">
	<input type="submit" name="submit" value="读 取" > 
	</td>
	</tr>
</table>
</form>
<%
End Sub
Sub doout()
		Dim Rs,node,skid,cssid,RsSkin,i
		If Request.form("skid")="" Then
				Head()
			Errmsg=ErrMsg + "<BR><li>必须选择要导出的模板."
			dvbbs_error()
			Call Footer()	
			Exit Sub
		End If
	 If SkinConnection(Request.Form("skinmdb")) Then
		Set Rs=Dvbbs.Execute("select Forum_Css From Dv_Setup")
		Set CssList=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		CssList.LoadXMl Rs(0)
		Set Rs=Nothing
		For Each skid in Request.form("skid") 
			Set Rs=Dvbbs.Execute("select * From Dv_style Where ID="&CLng(skid))
			Set cssid=Request.Form("cssid_"&skid)
			Set RsSkin=Server.CreateObject("adodb.recordset")
			RsSkin.open "Select * From Dv_Style Where Id=0",StyleConn,1,3
			RsSkin.AddNew
			RsSkin("StyleName")=Request.Form("StyleName"&skid)
			RsSkin("Main_Style")=Rs("Main_Style")
			RsSkin("Style_Pic")=Rs("Style_Pic")
			RsSkin("page_index")=Rs("page_index")
			RsSkin("page_dispbbs")=Rs("page_dispbbs")
			RsSkin("page_showerr")=Rs("page_showerr")
			RsSkin("page_login")=Rs("page_login")
			RsSkin("page_online")=Rs("page_online")
			RsSkin("page_usermanager")=Rs("page_usermanager")
			RsSkin("page_fmanage")=Rs("page_fmanage")
			RsSkin("page_boardstat")=Rs("page_boardstat")
			RsSkin("page_paper_even_toplist")=Rs("page_paper_even_toplist")
			RsSkin("page_query")=Rs("page_query")
			RsSkin("page_show")=Rs("page_show")
			RsSkin("page_dispuser")=Rs("page_dispuser")
			RsSkin("page_help_permission")=Rs("page_help_permission")
			RsSkin("page_postjob")=Rs("page_postjob")
			RsSkin("page_post")=Rs("page_post")
			RsSkin("page_boardhelp")=Rs("page_boardhelp")
			RsSkin("Forum_CSS")=OutCSSDom(cssid).xml
			RsSkin("DateAndTime")=Now()
			RsSkin("Readme")=Request.Form("readme"&skid)
			RsSkin.Update
			RsSkin.Close
		Next
		Head()
		Dv_suc "模板数据已经保存到您的论坛根目录下的skins下,文件名为"&Request.Form("skinmdb") 
		Call Footer()	
	 Else
	 		Head()
			Errmsg=ErrMsg + "<BR><li>目标数据库"& Request.Form("skinmdb")&"不存在或有错误.</li>"
			dvbbs_error()
			Call Footer()	
	 End If
End Sub
Function OutCSSDom(IDlist)
	Dim XML
	Set XML=CssList.cloneNode(True)
	Dim Node,cssid,id
	cssid=""
		For Each id in IDlist
			If IsNumeric(id) and id<>"" Then
				If cssid="" Then
					cssid="@id !="&id&" "
				Else
					cssid=cssid & "and @id !="&Id&" "
				End If
			End If
		Next
		If CssID<>"" Then CssID="["&CssID&"]"
		For Each Node in XML.documentElement.SelectNodes("css"&CSSID)
			XML.documentElement.removeChild(node)
		Next
		Set OutCSSDom=XML
End Function
Function  SkinConnection(mdbname)
	On Error Resume Next 
	Set StyleConn = Server.CreateObject("ADODB.Connection")
	StyleConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(MyDbPath &"skins/"&mdbname)
	If Err  Then 
		err.Clear
		SkinConnection=False
	Else
			SkinConnection=True
	End If
End Function
Sub Readme()
	%>
<table border="0" cellspacing="1" cellpadding="5" align=center width="95%" class="tableBorder">
<tr>
<th colspan="3" align="center" id="TableTitleLink">
<a href="Loadskin.asp">Dvbbs 7.1.0 Sp1 模板和CSS导入和导出管理</a>
</th>
</tr>
<tr>
<td class="forumHeaderBackgroundAlternate">请选择操作：
<a href="Loadskin.asp" style="color : blue;">导 出</a> 
| <a href="Loadskin.asp?action=load" style="color : blue;">导 入</a>
</td>
</tr>
</table>
<br>
	<%
End Sub
Sub Skinlist()
	Dim Rs,node
	Set Rs=Dvbbs.Execute("select Forum_Css From Dv_Setup")
	Set CssList=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	CssList.LoadXMl Rs(0)
	Set Rs=Nothing
	Set Rs=Dvbbs.Execute("select ID,StyleName From Dv_style")
	%>
	<form method="post" action="Loadskin.asp?action=doout">
<table border="0" cellspacing="1" cellpadding="4" align="center" width="95%" class="tableBorder">
	<tr><th colspan="4">模板和CSS的导出</th>
	</tr>
		<tr>
		<td class="forumrow" align="left" colspan="4" style="padding:10px;">您现在要执行的是模板和CSS的导出,把需要导出的勾上,设置好保存的数据库名称,然后提交.</td>
			</tr>
	<tr>
	
	<td width="40"  align="center" class="bodytitle">选择</td>
	<td width="100"  align="center" class="bodytitle">模版名称</td>
	<td width="150" class="bodytitle">附加说明</td>
	<td width="402"  class="bodytitle">适用的CSS样式</td>
	</tr>
	<%
	do while not Rs.eof
	%>
	<tr>
	<td width="40" class="forumrow" align="center"><input type="checkbox" name="skid" value="<%=Rs("id")%>"></td>
	<td width="100" class="forumrow" align="center"><Input tyle="text" Name="StyleName<%=Rs("id")%>" value="<%=Rs("StyleName")%>"></td>
	<td width="150" class="forumrow"><div align="left">
	<textarea name="readme<%=Rs("id")%>" rows="4" cols="30">Skin for <%=fversion%>(From <%=Dvbbs.Forum_info(0)%>)</textarea>
	</div></td>
	<td width="402" class="forumrow">
	<%
	For Each Node in CssList.documentElement.selectNodes("css[tid='"& Rs("id") &"']")
	%>
	<input type="checkbox" name="cssid_<%=Rs("id")%>" value="<%=node.selectSingleNode("@id").text%>">
	<%=node.selectSingleNode("@type").text%>
	<%
	Next
	%>
	</td>
	</tr>
	<%	Rs.movenext
		loop
		Rs.close:Set Rs=Nothing
	%>
	<tr>
	<td colspan="4" align=center class="forumRowHighlight">
	导出到数据库：<input type="text" name="skinmdb" size="30" value="Dv_skin.mdb">
	<input type="submit" name="submit" value="导 出" > 
	</td>
	</tr>
	</table>
	</form>
	<%
End Sub
%>