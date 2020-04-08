<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/dv_clsother.asp" -->
<!-- #include File="inc/Upload_Class.asp" -->
<%
Const UploadFileType = "gif|jpg|jpeg|bmp|png|rar|txt|zip|mid"

Dim groupboardid
groupboardid = Cint(Request("groupboardid"))

If Request("t")="1" Then
	Upfile_Main()
Else
	Main()
End If
Dvbbs.PageEnd()
Sub Main()
	Dvbbs.Loadtemplates("")
	Dvbbs.Head()
	Dvbbs.ShowErr
	Dim dateupnum
	Dim upset
	upset=CInt(Dvbbs.Groupsetting(66))
	dateupnum=Cint(Dvbbs.UserToday(2))
%>
<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<table border=0 cellspacing=0 cellpadding=0 style="width:100%;height:100%">
<form name="form" method="post" action="IndivGroup_Upload.asp?t=1&groupboardid=<%=groupboardid%>" enctype="multipart/form-data">
<tr>
<%
If Cint(Dvbbs.Groupsetting(7))=0 Then
	Response.Write "<TD class=tablebody2 valign=top>您没有在本论坛上传文件的权限!</td>"
Else
	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
%>
<INPUT TYPE="hidden" NAME="UploadCode" value="<%=PostRanNum%>">
<TD id="upid" class=tablebody1 valign=top>
	<input type="file" name="file1" width=200 value="" size="40" /></TD>
<td class=tablebody1 valign=top width=1>
	<input type="submit" name="Submit" value="上传" onclick="parent.document.PostForm.submit.disabled=true;">
</td>
<td class=tablebody1 valign=top>
	<font color=<%=Dvbbs.mainsetting(1) %> >今天还可上传<%=Dvbbs.Groupsetting(50)-dateupnum%>个</font>；
	<a style="CURSOR: help" title="论坛限制:一次<%=Dvbbs.Groupsetting(40)%>个，一天<%=Dvbbs.Groupsetting(50)%>个,每个<%=Dvbbs.Groupsetting(44)%>K">(查看论坛限制)</a>
</TD>
<%end if%>
</tr>
</form>
</table>
</body>
</html>
<%
End Sub

Sub Upfile_Main()
	Dvbbs.Loadtemplates("")
	Dvbbs.Head()
	Dvbbs.ShowErr()
	Server.ScriptTimeOut=999999'要是你的论坛支持上传的文件比较大，就必须设置。
	'-----------------------------------------------------------------------------
	'提交验证
	'-----------------------------------------------------------------------------
	If Not Dvbbs.ChkPost Then
		Response.End
	End If
	If Dvbbs.Userid=0 Then
		Response.write "你还未登陆！"
		Response.End
	End If
	If Cint(Dvbbs.GroupSetting(7))=0 then
		Response.write "您没有在本论坛上传文件的权限"
		Response.End
	End If
%>
	<script>
	parent.document.PostForm.submit.disabled=false;
	</script>
	<table width="100%" height="100%" border=0 cellspacing=0 cellpadding=0>
	<tr><td class=tablebody2 valign=top height=40>
	<%
	UploadFile
	%>
	</td></tr>
	</table>
</body>
</html>
<%
End Sub

Sub UploadFile()
	Dim Upload,FormName,FilePath,ChildFilePath
	Dim File,F_FileName,F_ViewName,F_Filesize,F_FileExt,F_Type
	Dim Previewpath,DrawInfo,InceptMaxFile
	Dim OnceUPCount
	OnceUPCount = Request.Cookies("upNum")
	If OnceUPCount = "" or Not Isnumeric(OnceUPCount) Then
		OnceUPCount = 0
	Else
		OnceUPCount = Clng(OnceUPCount)
	End If
	If OnceUPCount >= Clng(Dvbbs.GroupSetting(40)) then
 		Response.write "一次只能上传"&Dvbbs.GroupSetting(40)&"个文件！"
		Exit Sub
	Else
		InceptMaxFile = Clng(Dvbbs.GroupSetting(40)) - OnceUPCount
	End If
	If Not IsNumeric(Dvbbs.UserToday(2)) Then Dvbbs.UserToday(2) = 0
	If Clng(Dvbbs.UserToday(2))>Clng(Dvbbs.GroupSetting(50)) Then
 		Response.write "已超出了你在论坛每天上传的文件个数"&Dvbbs.GroupSetting(50)&"个！"
		Exit Sub
	Else
		If Clng(Dvbbs.GroupSetting(50))-Clng(Dvbbs.UserToday(2))<InceptMaxFile Then
			InceptMaxFile = Clng(Dvbbs.GroupSetting(50))-Clng(Dvbbs.UserToday(2))
		End If
	End If

	'上传目录
	FilePath = CreatePath(CheckFolder)
	'不带系统上传目录的下级目录路径
	ChildFilePath = Replace(FilePath,CheckFolder,"")
	'预览图片目录路径
	Previewpath = "PreviewImage/"
	Previewpath = CreatePath(Previewpath)

	If Dvbbs.Forum_UploadSetting(17)="1" Then
		DrawInfo = Dvbbs.Forum_UploadSetting(4)
	ElseIf Dvbbs.Forum_UploadSetting(17)="2" Then
		DrawInfo = Dvbbs.Forum_UploadSetting(9)
	Else
		DrawInfo = ""
	End If
	If DrawInfo = "0" Then
		DrawInfo = ""
		Dvbbs.Forum_UploadSetting(17) = 0
	End If

	Set Upload = New UpFile_Cls
	Upload.UploadType			= Cint(Dvbbs.Forum_UploadSetting(2))	'设置上传组件类型
	Upload.UploadPath			= FilePath								'设置上传路径
	Upload.InceptFileType		= Replace(UploadFileType,"|",",")		'设置上传文件限制
	Upload.MaxSize				= Int(Dvbbs.GroupSetting(44))			'单位 KB
	Upload.InceptMaxFile		= InceptMaxFile							'每次上传文件个数上限
	Upload.ChkSessionName		= "UploadCode"							'防止重复提交，SESSION名与提交的表单要一致。
	'预览图片设置
	Upload.PreviewType			= Cint(Dvbbs.Forum_UploadSetting(3))	'设置预览图片组件类型
	Upload.PreviewImageWidth	= Dvbbs.Forum_UploadSetting(14)			'设置预览图片宽度
	Upload.PreviewImageHeight	= Dvbbs.Forum_UploadSetting(15)			'设置预览图片高度
	Upload.DrawImageWidth		= Dvbbs.Forum_UploadSetting(11)			'设置水印图片或文字区域宽度
	Upload.DrawImageHeight		= Dvbbs.Forum_UploadSetting(12)			'设置水印图片或文字区域高度
	Upload.DrawGraph			= Dvbbs.Forum_UploadSetting(10)			'设置水印透明度
	Upload.DrawFontColor		= Dvbbs.Forum_UploadSetting(6)			'设置水印文字颜色
	Upload.DrawFontFamily		= Dvbbs.Forum_UploadSetting(7)			'设置水印文字字体格式
	Upload.DrawFontSize			= Dvbbs.Forum_UploadSetting(5)			'设置水印文字字体大小
	Upload.DrawFontBold			= Dvbbs.Forum_UploadSetting(8)			'设置水印文字是否粗体
	Upload.DrawInfo				= DrawInfo								'设置水印文字信息或图片信息
	Upload.DrawType				= Dvbbs.Forum_UploadSetting(17)			'0=不加载水印 ，1=加载水印文字，2=加载水印图片
	Upload.DrawXYType			= Dvbbs.Forum_UploadSetting(13)			'"0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
	Upload.DrawSizeType			= Dvbbs.Forum_UploadSetting(16)			'"0"=固定缩小，"1"=等比例缩小
	If Dvbbs.Forum_UploadSetting(18)<>"" or Dvbbs.Forum_UploadSetting(18)<>"0" Then
		Upload.TransitionColor	= Dvbbs.Forum_UploadSetting(18)			'透明度颜色设置
	End If
	'执行上传
	Upload.SaveUpFile
	If Upload.ErrCodes<>0 Then
		Response.write "错误："& Upload.Description & "[ <a href=""IndivGroup_Upload.asp?groupboardid=" & groupboardid & """>重新上传</a> ]"
		Exit Sub
	End If
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
				F_FileName = FilePath & File.FileName
				'创建预览及水印图片
				If Upload.PreviewType<>999 and File.FileType=1 then
						F_Viewname = Previewpath & "pre" & Replace(File.FileName,File.FileExt,"") & "jpg"
						'创建预览图片:Call CreateView(原始文件的路径,预览文件名及路径,原文件后缀)
						Upload.CreateView F_FileName,F_Viewname,File.FileExt
				End If
				UploadSave F_FileName,ChildFilePath&File.FileName,File.FileExt,F_Viewname,File.FileSize,File.FileType,File.FileOldName
			Set File = Nothing
		Next
	Else
		Response.write "请正确选择要上传的文件。[ <a href=""IndivGroup_Upload.asp?groupboardid=" & groupboardid & """>重新上传</a> ]"
		Exit Sub
	End If
	Call Suc_upload(Upload.Count,OnceUPCount)
	Set Upload = Nothing
End Sub

'保存上传数据并返回附件ID
Sub UploadSave(FileName,ChildFileName,FileExt,ViewName,FileSize,F_Type,F_OldName)
	Dim ShwoFileName
	ShwoFileName = Dvbbs.Checkstr(Replace(FileName,CheckFolder,"UploadFile/"))
	ChildFileName = Dvbbs.Checkstr(ChildFileName)
	Dvbbs.Execute("Insert into Dv_upFile (F_BoardID,F_UserID,F_Username,F_Filename,F_Viewname,F_FileType,F_Type,F_FileSize,F_Readme,F_Flag,F_OldName) values ("&groupboardid&","&Dvbbs.UserID&",'"&Dvbbs.Membername&"','"&ChildFileName&"','"&ViewName&"','"&FileExt&"',"&F_Type&","&FileSize&",'圈子上传图片',4,'"&Left(Dvbbs.Checkstr(F_OldName),50)&"')")
	Dim Rs,UpFileID,DownloadID
	Set Rs=Dvbbs.Execute("Select top 1 F_ID From Dv_upFile order by F_ID desc")
		DownloadID=rs(0)
		UpFileID = DownloadID & ","
	Rs.Close
	Set Rs=nothing
	If F_Type=1 or F_Type=2 then
		Response.write "<script>var str = '<br />[upload="&FileExt&","&F_OldName&"]"&ShwoFileName&"[/upload]<br>'</script>"
	Else
		Response.write "<script>var str ='<br />[upload="&FileExt&","&F_OldName&"]viewFile.asp?ID="&DownloadID&"[/upload]<br>'</script>"
	End If
	Response.write "<script>parent.dvtextarea.insert(str);</script>"
End Sub

Sub Suc_upload(UpCount,upNum)
	upNum = upNum + UpCount
	Response.Cookies("upNum") = upNum
	Dim iUserInfo
	Dvbbs.UserToday(2) = Dvbbs.UserToday(2)+UpCount
	iUserInfo = Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Dvbbs.UserToday(2) & "|" & Dvbbs.UserToday(3) & "|" & Dvbbs.UserToday(4)
	iUserInfo=Dvbbs.Checkstr(iUserInfo)
	If upNum < Clng(Dvbbs.GroupSetting(40)) And Dvbbs.UserToday(2) < Clng(Dvbbs.GroupSetting(50)) Then
		Response.Write UpCount & "个文件上传成功,目前今天总共上传了" & Dvbbs.UserToday(2) & "个附件 [ <a href=IndivGroup_Upload.asp?groupboardid=" & groupboardid & ">继续上传</a> ]"
	Else
		Response.write UpCount & "个文件上传成功!本次已达到上传数上限。"
	End If
	Dvbbs.Execute("UPDATE [Dv_user] SET UserToday = '" & iUserInfo &"' WHERE UserID = " & Dvbbs.UserID)
	Dvbbs.UserSession.documentElement.selectSingleNode("userinfo/@usertoday").text=iUserInfo
End Sub


'读取上传目录
Function CheckFolder()
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	CheckFolder = Replace(Replace(Dvbbs.Forum_Setting(76),Chr(0),""),".","")
	'在目录后加(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'按月份自动明名上传文件夹,需要ＦＳＯ组件支持。
Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	'以年月创建上传文件夹，格式：2003－8
	uploadpath = year(now) & "-" & month(now)
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 Then
			CreatePath = PathValue & uploadpath & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function
%>