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
'�ϴ����ѡ��:Upload_type=����
Upload_type = Cint(Dvbbs.Forum_UploadSetting(2))

DvbbsWap.ShowXMLStar
ChkUpfile
If Not FoundErr Then
	If Upload_type=999 Then
		DvbbsWap.ShowErr 0,"ϵͳ�ѹر��ϴ������Ĺ��ܣ�"
	Else
		Upload_Main
	End If
End If
DvbbsWap.ShowXMLEnd

'��֤�û��ϴ�Ȩ��
Sub ChkUpfile()
	If Dvbbs.UserID = 0 Then
		DvbbsWap.ShowErr 0,"ֻ������̳�û������ϴ�������"
		FoundErr = True
		Exit Sub
	End If
	If Dvbbs.GroupSetting(7) = "0" Then
		DvbbsWap.ShowErr 0,"��û���ϴ�������Ȩ�ޣ�"
		FoundErr = True
		Exit Sub
	End If
	If Clng(Dvbbs.UserToday(2))>=Clng(Dvbbs.GroupSetting(50)) Then
		DvbbsWap.ShowErr 0,"ϵͳ���ƻ�Աÿ��ֻ���ϴ�"&Dvbbs.GroupSetting(50)&"��������"
		FoundErr = True
		Exit Sub
	End If
	If Request("t")="1" Then
		DvbbsWap.ShowErr 1,"ϵͳ֧���ϴ�������"
		FoundErr = True
		Exit Sub
	End If
End Sub

Sub Upload_Main()
Dim FormPath,Upload,FormName,File,F_FileName
Dim TempData
FormPath=CheckFolder&CreatePath()	'�ϴ�Ŀ¼·��
Set Upload = New UpFile_Cls
	Upload.UploadType			= Upload_type							'�����ϴ��������
	Upload.UploadPath			= FormPath								'�����ϴ�·��
	Upload.InceptFileType		= "gif,jpg,bmp,jpeg,png"		'�����ϴ��ļ�����
	Upload.MaxSize				= Int(Dvbbs.GroupSetting(44))			'��λ KB
	Upload.InceptMaxFile		= 1							'ÿ���ϴ��ļ���������
	'ִ���ϴ�
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
		DvbbsWap.ShowErr 0,"����ȷѡ��Ҫ�ϴ����ļ���[ �����ϴ� ]"
		Exit Sub
	End If
	If Upload.ErrCodes<>0 Then
		DvbbsWap.ShowErr 0,"����"& Upload.Description & "[ �����ϴ� ]"
		Exit Sub
	End If
	TempData = Dvbbs.UserToday(0) & "|" & Dvbbs.UserToday(1) & "|" & Clng(Dvbbs.UserToday(2))+Upload.Count &"|"& Dvbbs.UserToday(3) &"|"& Dvbbs.UserToday(4)
	Set Upload = Nothing
	Dvbbs.Execute("UPDATE [Dv_user] SET UserToday = '" & Dvbbs.Checkstr(TempData) &"' WHERE UserID = " & Dvbbs.UserID)
	DvbbsWap.ShowErr 1,"�ϴ��ɹ���"
End Sub

'��ȡ�ϴ�Ŀ¼
Function CheckFolder()
	If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
	CheckFolder = Replace(Replace(Dvbbs.Forum_Setting(76),Chr(0),""),".","")
	'��Ŀ¼���(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'���·��Զ������ϴ��ļ���,��Ҫ�ƣӣ����֧�֡�
Private Function CreatePath()
	Dim objFSO,Fsofolder,uploadpath
	uploadpath=year(now)&"-"&month(now)	'�����´����ϴ��ļ��У���ʽ��2003��8
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
