<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include file="inc/Class_Mobile.asp"-->
<!-- #include File="inc/Upload_Class.asp" -->
<%
Dim FoundErr
Dim Upload_type,Forum_Url
Forum_Url = Dvbbs.Get_ScriptNameUrl
FoundErr = False
'---------------------------------------------------------------
'上传组件选择:Upload_type=参数
Upload_type = Cint(Dvbbs.Forum_UploadSetting(2))

DvbbsWap.ShowXMLStar
ChkUpfile
If Not FoundErr Then
	If Upload_type=999 Then
		DvbbsWap.ShowErr 0,"系统已关闭上传附件的功能！"
	Else
		Upload_Main
	End If
End If
DvbbsWap.ShowXMLEnd

'验证用户上传权限
Sub ChkUpfile()
	If Dvbbs.UserID = 0 Then
		DvbbsWap.ShowErr 0,"只允许论坛用户才能上传附件！"
		FoundErr = True
		Exit Sub
	End If
	If Dvbbs.GroupSetting(7) = "0" Then
		DvbbsWap.ShowErr 0,"你没有上传附件的权限！"
		FoundErr = True
		Exit Sub
	End If
	If Clng(Dvbbs.UserToday(2))>=Clng(Dvbbs.GroupSetting(50)) Then
		DvbbsWap.ShowErr 0,"系统限制会员每天只能上传"&Dvbbs.GroupSetting(50)&"个附件！"
		FoundErr = True
		Exit Sub
	End If
	If Request("t")="1" Then
		DvbbsWap.ShowErr 1,"系统支持上传附件。"
		FoundErr = True
		Exit Sub
	End If
End Sub

Sub Upload_Main()
Dim FormPath,Upload,FormName,File,F_FileName
Dim TempData
FormPath=CheckFolder&CreatePath()	'上传目录路径
Set Upload = New UpFile_Cls
	Upload.UploadType			= Upload_type							'设置上传组件类型
	Upload.UploadPath			= FormPath								'设置上传路径
	Upload.InceptFileType		= "gif,jpg,bmp,jpeg,png"		'设置上传文件限制
	Upload.MaxSize				= Int(Dvbbs.GroupSetting(44))			'单位 KB
	Upload.InceptMaxFile		= 1							'每次上传文件个数上限
	'执行上传
	Upload.SaveUpFile
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
				F_FileName = FormPath & File.FileName
				Response.Write "<fileurl>"
				Response.Write Forum_Url &"/"& F_FileName
				Response.Write "</fileurl>"
			Set File = Nothing
		Next
	Else
		DvbbsWap.ShowErr 0,"请正确选择要上传的文件。[ 重新上传 ]"
		Exit Sub
	End If
	If Upload.ErrCodes<>0 Then
		DvbbsWap.ShowErr 0,"错误："& Upload.Description & "[ 重新上传 ]"
		Exit Sub
	End If
	TempData = Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Clng(Dvbbs.UserToday(2))+Upload.Count &"|"& Dvbbs.UserToday(3) &"|"& Dvbbs.UserToday(4)
	Set Upload = Nothing
	Dvbbs.Execute("UPDATE [Dv_user] SET UserToday = '" & Dvbbs.Checkstr(TempData) &"' WHERE UserID = " & Dvbbs.UserID)
	DvbbsWap.ShowErr 1,"上传成功！"
End Sub

'读取上传目录
Function CheckFolder()
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	CheckFolder = Replace(Replace(Dvbbs.Forum_Setting(76),Chr(0),""),".","")
	'在目录后加(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'按月份自动明名上传文件夹,需要ＦＳＯ组件支持。
Private Function CreatePath()
	Dim objFSO,Fsofolder,uploadpath
	uploadpath=year(now)&"-"&month(now)	'以年月创建上传文件夹，格式：2003－8
	Fsofolder = Server.MapPath(CheckFolder & uploadpath)
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Fsofolder) = False Then
			objFSO.CreateFolder Fsofolder
		End If
		If Err.Number = 0 Then
			CreatePath = uploadpath & "/"
		Else
			CreatePath = ""
		End If
	Set objFSO = Nothing
End Function
%>
