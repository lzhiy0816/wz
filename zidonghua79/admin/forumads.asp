<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/dv_clsother.asp" -->
<%
Call Head()
Dim admin_flag
admin_flag=",2,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>本页面为管理员专用，请<a href=../admin_login.asp target=_top>登录</a>后进入。<br><li>您没有管理本页面的权限。"
	dvbbs_error()
Else 
	If request("action")="save" Then 
		Call saveconst()
	Else
		Call consted()
	End If
		If founderr then call dvbbs_error()
		footer()
	End If 

Sub consted()
dim sel
%>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">

<tr> 
<th height="23" colspan="2" class="tableHeaderText"><b>论坛广告设置</b>（如为设置分论坛，就是分论坛首页广告，下属页面为帖子显示页面）</th>
</tr>
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2><B>说明</B>：<BR>1、复选框中选择的为当前的使用设置模板，点击可查看该模板设置，点击别的模板直接查看该模板并修改设置。您可以将您下面的设置保存在多个论坛版面中<BR>2、您也可以将下面设定的信息保存并应用到具体的分论坛版面设置中，可多选<BR>3、如果您想在一个版面引用别的版面的配置，只要点击该版面名称，保存的时候选择要保存到的版面名称名称即可。
<hr size=1 width="100%" color=blue>
</td>
</tr>
<FORM METHOD=POST ACTION="">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
查看分版面广告设置，请选择左边下拉框相应版面&nbsp;&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value="">查看分版面广告请选择</option>
<%
Dim ii
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value=""forumads.asp?boardid="&rs(0)&""">"
Select Case rs(2)
	Case 0
		Response.Write "╋"
	Case 1
		Response.Write "&nbsp;&nbsp;├"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;│"
	Next
	Response.Write "&nbsp;&nbsp;├"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
</FORM>
</table><BR>
<form method="POST" action=forumads.asp?action=save>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
<input type=checkbox name="getskinid" value="1" <%if request("getskinid")="1" or request("boardid")="" then Response.Write "checked"%>><a href="forumads.asp?getskinid=1">论坛默认广告</a><BR> 点击此处返回论坛默认广告设置，默认广告设置包含所有<FONT COLOR="blue">除</FONT>包含具体版面内容（如帖子列表、帖子显示、版面精华、版面发贴等）<FONT COLOR="blue">以外</FONT>的页面。<hr size=1 width="90%" color=blue>
</td>
</tr>
<tr> 
<td width="200" class="forumrow" valign=top>
版面广告保存选项<BR>
请按 CTRL 键多选<BR>
<select name="getboard" size="28" style="width:100%" multiple>
<%
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value="&rs(0)&">"
Select Case rs(2)
	Case 0
		Response.Write "╋"
	Case 1
		Response.Write "&nbsp;&nbsp;├"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;│"
	Next
	Response.Write "&nbsp;&nbsp;├"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
<td class="forumrow" valign=top>
<table>
<tr>
<td width="200" class="forumrow"><B>首页顶部广告代码</B><BR>如果开启了互动广告功能中的顶部广告，此处设置为无效</td>
<td width="*" class="forumrow"> 
<textarea name="Forum_ads(0)" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(0))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>首页尾部广告代码</B></td>
<td width="*" class="forumrow"> 
<textarea name="Forum_ads(1)" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(1))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>开启首页浮动广告</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(2)" value=0 <%if Dvbbs.Forum_ads(2)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_ads(2)" value=1 <%if Dvbbs.Forum_ads(2)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页浮动广告图片地址</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(3)" size="35" value="<%=Dvbbs.Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页浮动广告连接地址</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(4)" size="35" value="<%=Dvbbs.Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页浮动广告图片宽度</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(5)" size="3" value="<%=Dvbbs.Forum_ads(5)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页浮动广告图片高度</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(6)" size="3" value="<%=Dvbbs.Forum_ads(6)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>开启首页右下固定广告</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(13)" value=0 <%if Dvbbs.Forum_ads(13)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_ads(13)" value=1 <%if Dvbbs.Forum_ads(13)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页右下固定广告图片地址</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(8)" size="35" value="<%=Dvbbs.Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页右下固定广告连接地址</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(9)" size="35" value="<%=Dvbbs.Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页右下固定广告图片宽度</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(10)" size="3" value="<%=Dvbbs.Forum_ads(10)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>论坛首页右下固定广告图片高度</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(11)" size="3" value="<%=Dvbbs.Forum_ads(11)%>">&nbsp;象素
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>是否开启帖间随机广告</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(7)" value=0 <%if Dvbbs.Forum_ads(7)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_ads(7)" value=1 <%if Dvbbs.Forum_ads(7)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr> 
<td width="*" class="forumrow" valign="top" colspan=2><B>论坛帖间随机广告代码</B> <br>支持HTML语法，每条随机广告一行，用回车分开。</td>
</tr>
<tr>
<%
Dim Ads_14
If UBound(Dvbbs.Forum_ads)>13 Then
	Ads_14=Dvbbs.Forum_ads(14)
End If
%>
<td width="*" class="forumrow" colspan=2> 
<textarea name="Forum_ads(14)" style="width:100%" rows="10"><%=Ads_14%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>是否开启页面文字广告位</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(12)" value=0 <%if Dvbbs.Forum_ads(12)="0" then%>checked<%end if%>>关闭&nbsp;
<input type=radio name="Forum_ads(12)" value=1 <%if Dvbbs.Forum_ads(12)="1" then%>checked<%end if%>>打开&nbsp;
</td>
</tr>
<tr>
<%
Dim Ads_15
If UBound(Dvbbs.Forum_ads)>14 Then
	Ads_15=Dvbbs.Forum_ads(15)
End If
%>
<td width="200" class="forumrow"><B>页面文字广告位设置(版面)</B><BR>请确认已打开了页面文字广告位功能<BR></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(15)" value=0 <%if Ads_15="0" then%>checked<%end if%>>帖子列表&nbsp;
<input type=radio name="Forum_ads(15)" value=1 <%if Ads_15="1" then%>checked<%end if%>>帖子内容&nbsp;
<input type=radio name="Forum_ads(15)" value=2 <%if Ads_15="2" then%>checked<%end if%>>两者都显示&nbsp;
<input type=radio name="Forum_ads(15)" value=3 <%if Ads_15="3" then%>checked<%end if%>>两者都不显示&nbsp;
</td>
</tr>
<tr>
<%
Dim Ads_17
If UBound(Dvbbs.Forum_ads)>16 Then
	Ads_17=Dvbbs.Forum_ads(17)
End If
%>
<td width="200" class="forumrow"><B>文字广告每行广告个数</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(17)" size="3" value="<%=Ads_17%>">&nbsp;个
</td>
</tr>
<tr> 
<td width="*" class="forumrow" valign="top" colspan=2><B>页面文字广告位内容</B> <br>支持HTML语法，每条广告一行，用回车分开。</td>
</tr>
<tr>
<td width="*" class="forumrow" colspan=2>
<%
Dim Ads_16
If UBound(Dvbbs.Forum_ads)>15 Then
	Ads_16=Dvbbs.Forum_ads(16)
End If
%>
<textarea name="Forum_ads(16)" style="width:100%" rows="10"><%=Ads_16%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow">&nbsp;</td>
<td width="*" class="forumrow"> 
<div align="center"> 
<input type="submit" name="Submit" value="提 交">
</div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
<%
end sub

Sub SaveConst()
	Dim iSetting
	For i = 0 To 30
		If Trim(Request.Form("Forum_ads("&i&")"))="" Then
			If i = 1 Or i = 0 Then
				iSetting = ""
			ElseIf i = 17 Then
				iSetting = 1
			Else
				iSetting = 0
			End If
		Else
			iSetting=Replace(Trim(Request.Form("Forum_ads("&i&")")),"$","")
		End If

		If i = 0 Then
			Dvbbs.Forum_ads = iSetting
		Else
			Dvbbs.Forum_ads = Dvbbs.Forum_ads & "$" & iSetting
		End If
	Next

	If Request("getskinid")="1" Then
		Sql = "Update Dv_Setup Set Forum_ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"'"
		Dvbbs.Execute(sql)
		Dvbbs.loadSetup()
	End If
	If Request("getboard")<>"" Then
		Sql = "Update Dv_Board Set Board_Ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"' Where BoardID In ("&Request("getboard")&")"
		Dvbbs.Execute(Sql)
		Dvbbs.loadSetup()
		RestoreBoardCache()
		Dvbbs.ReloadBoardCache request("getboard")
	Else
		Dvbbs.loadSetup()
		RestoreBoardCache()
	End If

	Dv_suc("广告设置成功！")
End Sub
Sub RestoreBoardCache()
	Dim Board,node
	Dvbbs. LoadBoardList()
	For Each node in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style/@id")
		Application.Contents.Remove(Dvbbs.CacheName & "_showtextads_"&node.text)
		For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
			Dvbbs.LoadBoardData board.text
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text)
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text&"_-time")
		Next
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text)
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text&"_-time")
	Next
End Sub
%>