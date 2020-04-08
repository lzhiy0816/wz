<!--#Include File="../../inc/Dv_ClsMain.asp"-->
<%
If Session("flag")<> "" Then Dvbbs.Master = True
'UpUserFaceFolder 
'如是独立的虚拟目录,则要写成"/uploadFace"如果是论坛目录下的普通目录,则写成""
'Const UpUserFaceFolder=""
Const UpUserFaceFolder="/uploadFace/"
MyDbPath = "../"
If IsNumeric(Dvbbs.UserHidden) = 0 or Dvbbs.Userhidden = "" Then Dvbbs.UserHidden = 2
If IsNumeric(Dvbbs.UserID) = 0 Or Dvbbs.UserID="" Then Dvbbs.UserID=0
Dvbbs.UserID = Clng(Dvbbs.UserID)
Dvbbs.MemberClass = Dvbbs.checkStr(Request.Cookies(Dvbbs.Forum_sn)("userclass"))

Set MyBoardOnline=new Cls_UserOnlne 
'获得论坛基本信息和检测用户登陆状态
Dvbbs.GetForum_Setting
Dvbbs.CheckUserLogin
'重新赋予用户是否可进入后台权限		已移至Admin_login.asp验证，轻飘飘
'If Dvbbs.GroupSetting(70)="1" Then
'	Dvbbs.Master = True
'Else
'	Dvbbs.Master = False
'End If

'后台信息和函数部分
Dim AllPostTable
Dim AllPostTableName
Dim FoundErr
FoundErr=False 
Dim ErrMsg
Dim Rs,sql,i
Dvbbs.LoadTemplates("")
Set Rs=Dvbbs.Execute("Select H_Content From Dv_Help Where H_ID=1")
template.value = Rs(0)
'页面错误提示信息
Sub dvbbs_error()
	Response.Write"<br>"
	Response.Write"<table cellpadding=3 cellspacing=1 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>错误信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""Forumrow"" colspan=2>"
	Response.Write ErrMsg
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""Forumrow"" valign=middle colspan=2 align=center><a href=""javascript:history.go(-1)""><<返回上一页</a></td></tr>"
	Response.Write"</table>"
	footer()
	Response.End 
End Sub 
Function AllPostTable1()
	Dim Trs
	Set Trs=Dvbbs.Execute("select * from [Dv_TableList]")
	AllPostTable=""
	Do While Not TRs.EOF
		If AllPostTable=""  Then 
			AllPostTable=TRs("TableName")
			AllPostTableName=TRs("TableType")
		Else
			AllPostTable=AllPostTable&"|"&TRs("TableName")
			AllPostTableName=AllPostTableName&"|"&TRs("TableType")
		End If
	TRs.MoveNext
	Loop 
	Trs.Close 
	
End Function 
AllPostTable1
AllPostTableName=Split(AllPostTableName,"|")
AllPostTable=Split(AllPostTable,"|")
Dim NowUseBbs
NowUseBbs=Dvbbs.NowUseBbs

Sub footer()
	Response.Write"<table align=center >"
	Response.Write "<tr align=center><td width=""100%"" class=copyright>"
	Response.Write"Dvbbs v7.1 , Copyright (c) 2001-2005 <a href=""http://www.aspsky.net"" target=""_blank""><font color=#708796><b>AspSky<font color=#CC0000>.Net</font></b></font></a>. All Rights Reserved ."
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"</table>"
	Response.Write "</body>"
	Response.Write "</html>"
	SaveLog()
	Set Dvbbs=Nothing 
End Sub
Sub Head()
	Response.Write "<html>"
	Response.Write Chr(10)	
	Response.Write "<head>"
	Response.Write Chr(10)	
	Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
	Response.Write Chr(10)	
	Response.Write "<meta name=keywords content=""动网先锋,动网论坛,dvbbs"">"
	Response.Write Chr(10)	
	Response.Write "<meta name=""description"" content=""Design By www.dvbbs.net"">"
	Response.Write Chr(10)	
	Response.Write "<title>"& Dvbbs.Forum_info(0)&"-管理页面</title>"
	Response.Write Chr(10)	
	Response.Write Replace(template.html(1),"{$path}","images/")
	Response.Write Chr(10)
	Response.Write "</head>"
	Response.Write "<script src=""images/admin.js"" type=""text/javascript""></script><script src=""../inc/main.js"" type=""text/javascript""></script>"
	Response.Write Chr(10)
	Dim XMLDOM
	Set XMLDOM=Application(Dvbbs.CacheName&"_ssboardlist").cloneNode(True)
	Response.Write "<script language=""javascript"" type=""text/javascript"">"
	Response.Write "var boardxml='<?xml version=""1.0"" encoding=""gb2312""?>"& replace(XMLDom.documentElement.XML ,"'","\'")&"';"
	Response.Write "</script>"
	Response.Write "<body leftmargin=""0"" topmargin=""0"" marginheight=""0"" marginwidth=""0"">"
	Response.Write Chr(10)
%>
<div class=menuskin id=popmenu onmouseover="clearhidemenu();" onmouseout="dynamichide(event)" style="Z-index:100"></div>
<table cellpadding="3" cellspacing="0" border="0" align=center class="tableBorder1" style="width:100%">
<tr><td height=10></td></tr>
</table>
<%
End Sub
Sub Dv_suc(info)
	Response.Write"<br>"
	Response.Write"<table cellpadding=0 cellspacing=0 align=center class=""tableBorder"" style=""width:75%"">"
	Response.Write"<tr align=center>"
	Response.Write"<th width=""100%"" height=25 colspan=2>成功信息"
	Response.Write"</td>"
	Response.Write"</tr>"
	Response.Write"<tr>"
	Response.Write"<td width=""100%"" class=""forumRowHighlight"" colspan=2 height=25>"
	Response.Write info
	Response.Write"</td></tr>"
	Response.Write"<tr>"
	Response.Write"<td class=""forumRowHighlight"" valign=middle colspan=2 align=center><a href="&Request.ServerVariables("HTTP_REFERER")&" ><<返回上一页</a></td></tr>"
	Response.Write"</table>"
End Sub

Function fixjs(Str)
	If Str <>"" Then
		str = replace(str,"\", "\\")
		Str = replace(str, chr(34), "\""")
		Str = replace(str, chr(39),"\'")
		Str = Replace(str, chr(13), "\n")
		Str = Replace(str, chr(10), "\r")
		str = replace(str,"'", "&#39;")
	End If
	fixjs=Str
End Function
Function enfixjs(Str)
	If Str <>"" Then
		Str = replace(str,"&#39;", "'")
		Str = replace(str,"\""" , chr(34))
		Str = replace(str, "\'",chr(39))
		Str = Replace(str, "\r", chr(10))
		Str = Replace(str, "\n", chr(13))
		Str = replace(str,"\\", "\")
	End If
	enfixjs=Str
End Function

Function Reload_All_Board_Cache()
	'更新版面列表缓存
	ReloadBoardListAll
	'更新单个版面缓存（循环）
	Dim BoardListAll,BoardListNum,myBoardID
	BoardListAll=myCache.value
	BoardListNum=Ubound(BoardListAll,2)
	For i=0 To BoardListNum
		myBoardID=BoardListAll(0,i)
		ReloadBoardInfo(myBoardID)
		Set rs=Dvbbs.Execute("Select ParentStr from board where boardid="&myBoardID)
		If not rs.eof Then
			Dvbbs.ReloadBoardParentStr(rs(0))
		End If
		Rs.close
		Set Rs=nothing
	Next
End Function
Sub SaveLog()
	On Error Resume Next
	Dim RequestStr
	RequestStr= Request("action")
	If RequestStr<>"" Then 
		RequestStr="action="&RequestStr
		RequestStr=Dvbbs.checkStr(RequestStr)
		RequestStr=Left(RequestStr,250)
		sql="insert into [Dv_log] (l_touser,l_username,l_content,l_ip,l_type) values ('"&Dvbbs.ScriptName&"','"&Dvbbs.membername&"','"&RequestStr&"','"&Dvbbs.UserTrueIP&"',0)"		
		Dvbbs.Execute(sql)
	End If
	If request.form<>"" Then
		RequestStr=Dvbbs.checkStr(request.form)
		RequestStr=Left(RequestStr,250)
		sql="insert into [Dv_log] (l_touser,l_username,l_content,l_ip,l_type) values ('"&Dvbbs.ScriptName&"','"&Dvbbs.membername&"','"&RequestStr&"','"&Dvbbs.UserTrueIP&"',1)"		
		Dvbbs.Execute(sql)
	End If
End Sub
%>