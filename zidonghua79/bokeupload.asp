<!--#include FILE="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="boke/config.asp"-->
<!-- #include File="boke/Upload_Class.asp" -->
<%
DvBoke.Stats = "�ϴ�����"
DvBoke.Head(0)
Dim m
m = Request.QueryString("mode")
Select Case Request.QueryString("mode")
Case "UploadForm" : UploadForm()
Case "SaveUpload" : SaveUpload()
End Select

Sub UploadForm()
	Dim PageHtml
	DvBoke.LoadPage("topic.xslt")
	PageHtml = DvBoke.Page_Strings(26).text
	PageHtml = Replace(PageHtml,"{$upset}",5)
	PageHtml = Replace(PageHtml,"{$bokename}",Dvboke.BokeName)
	Response.Write PageHtml
End Sub

Sub SaveUpload()
	Dim MaxSize
	Server.ScriptTimeOut=999999'Ҫ�������̳֧���ϴ����ļ��Ƚϴ󣬾ͱ������á�
	'-----------------------------------------------------------------------------
	'�ύ��֤
	'-----------------------------------------------------------------------------
	If Not Dvbbs.ChkPost Then
		Response.End
	End If
	If Dvboke.Userid=0 Then
		Response.write "�㻹δ��½��"
		Response.End
	End If
	

	If Dvboke.System_UpSetting(0)="1" Then
		If Cint(Dvbbs.GroupSetting(7))=0 then
			Response.write "��û���ڱ���̳�ϴ��ļ���Ȩ�ޣ�"
			Response.End
		End If
	Else
		If Dvboke.System_UpSetting(1)="0" Then
			Response.write "��û���ڱ������ϴ��ļ���Ȩ�ޣ�"
			Response.End
		End If
	End If
	MaxSize = Int(Dvboke.System_UpSetting(1))
	If MaxSize=0 Then
		MaxSize = Int(Dvbbs.GroupSetting(44))
	End If
	If Dvboke.BokeNode.getAttribute("spacesize")=0 Then
		Response.write "��ǰ����������ϴ��ռ��������������Ա��ϵ��"
		Response.End
	ElseIf Dvboke.BokeNode.getAttribute("spacesize")>0 Then
		If Dvboke.BokeNode.getAttribute("spacesize")<MaxSize Then
			MaxSize = Dvboke.BokeNode.getAttribute("spacesize")
		End If
	End If

	'----------------------------------------------------------------------
	Dim FilePath,ChildFilePath,Previewpath,DrawInfo,AllFileName
	'�ϴ�Ŀ¼
	FilePath = CreatePath(CheckFolder)
	'����ϵͳ�ϴ�Ŀ¼���¼�Ŀ¼·��
	ChildFilePath = Replace(FilePath,CheckFolder,"")
	'Ԥ��ͼƬĿ¼·��
	Previewpath = "Boke/PreviewImage/"
	Previewpath = CreatePath(Previewpath)

	If Dvboke.System_UpSetting(17)="1" Then
		DrawInfo = Dvboke.System_UpSetting(4)
	ElseIf Dvboke.System_UpSetting(17)="2" Then
		DrawInfo = Dvboke.System_UpSetting(9)
	Else
		DrawInfo = ""
	End If
	If DrawInfo = "0" Then
		DrawInfo = ""
		Dvboke.System_UpSetting(17) = 0
	End If
%>
	<script>parent.document.PostForm.submit.disabled=false;</script>
<%
	Dim Upload,FormName,File,F_FileName,F_Viewname
	'========================================================================
	Set Upload = New UpFile_Cls
	Upload.UploadType			= Cint(Dvboke.System_UpSetting(2))	'�����ϴ��������
	Upload.UploadPath			= FilePath								'�����ϴ�·��
	Upload.InceptFileType		= Replace(DvBoke.System_Setting(16),"|",",")		'�����ϴ��ļ�����
	Upload.MaxSize				= MaxSize							'��λ KB
	Upload.InceptMaxFile		= 5									'ÿ���ϴ��ļ���������
	'Upload.ChkSessionName		= "UploadCode"							'��ֹ�ظ��ύ��SESSION�����ύ�ı�Ҫһ�¡�
	'Ԥ��ͼƬ����
	Upload.PreviewType			= Cint(Dvboke.System_UpSetting(3))	'����Ԥ��ͼƬ�������
	Upload.PreviewImageWidth	= Dvboke.System_UpSetting(14)			'����Ԥ��ͼƬ���
	Upload.PreviewImageHeight	= Dvboke.System_UpSetting(15)			'����Ԥ��ͼƬ�߶�
	Upload.DrawImageWidth		= Dvboke.System_UpSetting(11)			'����ˮӡͼƬ������������
	Upload.DrawImageHeight		= Dvboke.System_UpSetting(12)			'����ˮӡͼƬ����������߶�
	Upload.DrawGraph			= Dvboke.System_UpSetting(10)			'����ˮӡ͸����
	Upload.DrawFontColor		= Dvboke.System_UpSetting(6)			'����ˮӡ������ɫ
	Upload.DrawFontFamily		= Dvboke.System_UpSetting(7)			'����ˮӡ���������ʽ
	Upload.DrawFontSize			= Dvboke.System_UpSetting(5)			'����ˮӡ���������С
	Upload.DrawFontBold			= Dvboke.System_UpSetting(8)			'����ˮӡ�����Ƿ����
	Upload.DrawInfo				= DrawInfo								'����ˮӡ������Ϣ��ͼƬ��Ϣ
	Upload.DrawType				= Dvboke.System_UpSetting(17)			'0=������ˮӡ ��1=����ˮӡ���֣�2=����ˮӡͼƬ
	Upload.DrawXYType			= Dvboke.System_UpSetting(13)			'"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
	Upload.DrawSizeType			= Dvboke.System_UpSetting(16)			'"0"=�̶���С��"1"=�ȱ�����С
	If Dvboke.System_UpSetting(18)<>"" or Dvboke.System_UpSetting(18)<>"0" Then
		Upload.TransitionColor	= Dvboke.System_UpSetting(18)			'͸������ɫ����
	End If
	
	'ִ���ϴ�
	Upload.SaveUpFile
	
	If Upload.ErrCodes<>0 Then
		Response.write "����"& Upload.Description & "[ <a href=""?mode=UploadForm"">�����ϴ�</a> ]"
		Exit Sub
	End If
	If Upload.Count > 0 Then
		For Each FormName In Upload.UploadFiles
			Set File = Upload.UploadFiles(FormName)
				F_FileName = FilePath & File.FileName
				'Response.Write File.sFileName
				'REsponse.End
				'����Ԥ����ˮӡͼƬ
				If Upload.PreviewType<>999 and File.FileType=1 then
						F_Viewname = Previewpath & "pre" & Replace(File.FileName,File.FileExt,"") & "jpg"
						'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
						Upload.CreateView F_FileName,F_Viewname,File.FileExt
				End If
				UploadSave F_FileName,ChildFilePath&File.FileName,File.sFileName,File.FileExt,F_Viewname,File.FileSize,File.FileType
				AllFileName = AllFileName & "<br>" & Dvbbs.Get_ScriptNameUrl() & DvBoke.System_UpSetting(19) & ChildFilePath&File.FileName
			Set File = Nothing
		Next
	Else
		Response.write "����ȷѡ��Ҫ�ϴ����ļ���[ <a href=""?mode=UploadForm"">�����ϴ�</a> ]"
		Exit Sub
	End If

	Call Suc_upload(Upload.Count,Upload.CountSize,AllFileName)
	Set Upload = Nothing
End Sub

Sub Suc_upload(UpCount,upCountSize,AllFileName)
	Dim SucMsg
	Dim RemainSize
	upCountSize = Formatnumber((upCountSize/1024)/1024,2)
	IF Dvboke.BokeNode.getAttribute("spacesize")<>-1 Then
		RemainSize = cCur(Dvboke.BokeNode.getAttribute("spacesize")) - upCountSize
		If RemainSize<0 Then
			RemainSize = 0
		End If
	Else
		RemainSize = -1
	End If
	SucMsg = UpCount & "�������ϴ��ɹ�,��"&upCountSize&"MB ;"
	If RemainSize = 0 Then
		SucMsg = SucMsg & "[ ��ǰ���͵Ŀռ�����,�������Ա��ϵ��]"
	Else
		SucMsg = SucMsg & "[ <a href=""?mode=UploadForm"">�����ϴ�</a> ]"
	End If
	Dvboke.Execute("UPDATE [Dv_Boke_User] SET SpaceSize = '" & RemainSize &"' WHERE UserID = " & DvBoke.BokeUserID)
	Dim PageHtml
	DvBoke.LoadPage("topic.xslt")
	PageHtml = DvBoke.Page_Strings(27).text
	If m="1" Then SucMsg = SucMsg & "<br>�ļ���ַ��" & AllFileName
	PageHtml = Replace(PageHtml,"{$SucMsg}",SucMsg)
	Response.Write PageHtml
End Sub


'�����ϴ����ݲ����ظ���ID
Sub UploadSave(FileName,ChildFileName,sFileName,FileExt,ViewName,FileSize,F_Type)
'�����ֶ� ID,UserID,UserName,CatID,sType,TopicID,PostID,IsTopic,Title,FileName,FileType,FileSize,FileNote,DownNum,ViewNum,DateAndTime,PreviewImage 
'�ֶ����� ID=0 ,UserID=1 ,UserName=2 ,CatID=3 ,sType=4 ,TopicID=5 ,PostID=6 ,IsTopic=7 ,Title=8 ,FileName=9 ,FileType=10 ,FileSize=11 ,FileNote=12 ,DownNum=13 ,ViewNum=14 ,DateAndTime=15 ,PreviewImage=16 

'CatID,sType,TopicID,PostID,IsTopic,Title,FileNote,IsLock
	Dim ShwoFileName
	ShwoFileName = Dvbbs.Checkstr(Replace(FileName,CheckFolder,"UploadFile/"))
	ChildFileName = Dvbbs.Checkstr(ChildFileName)

	DvBoke.Execute("Insert into Dv_Boke_Upfile (BokeUserID,UserID,UserName,FileName,sFileName,FileType,FileSize,PreviewImage,IsLock,IsTopic) values ("& DvBoke.BokeUserID &","& DvBoke.UserID &",'"& DvBoke.BokeUserName &"','"& ChildFileName &"','"& Dvbbs.Checkstr(sFileName) &"',"& F_Type &","& FileSize &",'"& Dvbbs.Checkstr(ViewName) &"',4,-1)")

	Dim Rs,UpFileID,DownloadID
	Set Rs=DvBoke.Execute("Select top 1 ID From Dv_Boke_Upfile order by ID desc")
		DownloadID=rs(0)
		UpFileID = DownloadID & ","
	Rs.Close
	Set Rs=nothing
	'0=����,1=ͼƬ,2=FLASH,3=����,4=��Ӱ
	If F_Type=1 or F_Type=2 then
		Response.write "<script>parent.put_html('PostContent','[upload="&FileExt&"]"&ChildFileName&"[/upload]<br/>');</script>"
	Else
		Response.write "<script>parent.put_html('PostContent','[upload="&FileExt&"]BokeviewFile.asp?ID="&DownloadID&"[/upload]<br/>');</script>"
	End If
	Response.write "<script>parent.document.PostForm.upfilerename.value+='"&UpFileID&"'</script>"
End Sub


'��ȡ�ϴ�Ŀ¼
Function CheckFolder()
	If DvBoke.System_UpSetting(19)="" Or DvBoke.System_UpSetting(19)="0" Then DvBoke.System_UpSetting(19)="Boke/UploadFile/"
	CheckFolder = Replace(Replace(DvBoke.System_UpSetting(19),Chr(0),""),".","")
	'��Ŀ¼���(/)
	If Right(CheckFolder,1)<>"/" Then CheckFolder=CheckFolder&"/"
End Function

'���·��Զ������ϴ��ļ���,��Ҫ�ƣӣ����֧�֡�
Private Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	'�����´����ϴ��ļ��У���ʽ��2003��8
	uploadpath = year(now) & "-" & month(now)
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & DvBoke.BokeUserID ))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & DvBoke.BokeUserID)
		End If
		If objFSO.FolderExists(Server.MapPath(PathValue & DvBoke.BokeUserID &"/"& uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & DvBoke.BokeUserID &"/"& uploadpath)
		End If
		If Err.Number = 0 Then
			CreatePath = PathValue & DvBoke.BokeUserID &"/"& uploadpath &"/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function

%>