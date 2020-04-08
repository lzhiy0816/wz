<!--#include FILE="Upload.inc"-->
<%
'-----------------------------------------------------------------------
'--- �ϴ�������ģ��
'--- Copyright (c) 2004 Aspsky, Inc.
'--- Mail: Sunwin@artbbs.net   http://www.aspsky.net
'--- 2004-12-18
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- InceptFileType	: �����ϴ��������� (�Զ��ŷָ�����ļ�����) String
'-- MaxSize			: �����ϴ��ļ���С���� (��λ��kb) Long
'-- InceptMaxFile	: ����һ���ϴ��ļ������� Long
'-- UploadPath		: ���ñ����Ŀ¼���·�� String
'-- UploadType		: �����ϴ�������� ��0=������ϴ��࣬1=Aspupload3.0 ,2=SA-FileUp 4.0 ,3=DvFile.Upload V1.0��
'-- SaveUpFile		: ִ���ϴ�
'-- GetBinary		: �����ϴ��Ƿ񷵻��ļ�������  Bloonֵ : True/False
'-- ChkSessionName	: ����SESSION������ֹ�ظ��ύ��SESSION�����ύ�ı�����Ҫһ�¡�
'-- RName�����ļ���	: �����ļ���ǰ׺ (��Ĭ�����ɵ��ļ���Ϊ200412230402587123.jpg
'									���ã�RName="PRE_",���ɵ��ļ���Ϊ��PRE_200412230402587123.jpg)
'-----------------------------------------------------------------------
'-- ����ͼƬ�������
'-- PreviewType		: �������(0=CreatePreviewImage�����1=AspJpegV1.2 ,2=SoftArtisans ImgWriter V1.21)
'-- PreviewImageWidth	: ����Ԥ��ͼƬ����
'-- PreviewImageHeight	: ����Ԥ��ͼƬ�߶�
'-- DrawImageWidth	: ����ˮӡͼƬ�������������
'-- DrawImageHeight	: ����ˮӡͼƬ����������߶�
'-- DrawGraph		: ����ˮӡͼƬ����������͸����
'-- DrawFontColor	: ����ˮӡ������ɫ
'-- DrawFontFamily	: ����ˮӡ���������ʽ
'-- DrawFontSize	: ����ˮӡ���������С
'-- DrawFontBold	: ����ˮӡ�����Ƿ����
'-- DrawInfo		: ����ˮӡ������Ϣ��ͼƬ��Ϣ
'-- DrawType		: ���ü���ˮӡģʽ��0=������ˮӡ ��1=����ˮӡ���� ��2=����ˮӡͼƬ
'-- DrawXYType		: ͼƬ����ˮӡLOGOλ�����꣺"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
'-- DrawSizeType	: ����Ԥ��ͼƬ��С����"0"=�̶���С��"1"=�ȱ�����С
'-----------------------------------------------------------------------
'-- ��ȡ�ϴ���Ϣ
'-- ObjName			: ���õ��������
'-- Count			: �ϴ��ļ�����
'-- CountSize		: �ϴ��ܴ�С�ֽ���
'-- ErrCodes		: ����NUMBER (Ĭ��Ϊ0)
'-- Description		: ��������
'-----------------------------------------------------------------------
'-- CreateView Imagename,TempFilename,FileExt
'	����Ԥ��ͼƬ����: ԭʼ�ļ������·��,����Ԥ���ļ����·��,ԭ�ļ���׺
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- ��ȡ�ļ��������� : UploadFiles
'-- FormName		: ��������
'-- FileName		: ���ɵ��ļ�����
'-- FilePath		: �����ļ������·��
'-- FileSize		: �ļ���С
'-- FileContentType	: ContentType�ļ�����
'-- FileType		: 0=����,1=ͼƬ,2=FLASH,3=����,4=��Ӱ
'-- FileData		: �ļ������� (�������֧��ֱ�ӻ�ȡ���򷵻�Null)
'-- FileExt			: �ļ���׺
'-- FileWidth		: ͼƬ/Flash�ļ�����	�������ļ�Ĭ��=-1��
'-- FileHeight		: ͼƬ/Flash�ļ��߶�	�������ļ�Ĭ��=-1��
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-- ��ȡ������������ : UploadForms
'-- Count			: ������
'-- key				: ��������
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Class UpFile_Cls
	Private UploadObj,ImageObj
	Private FilePath,InceptFile,FileMaxSize,MaxFile,Upload_Type,FileInfo,IsBinary,SessionName
	Private Preview_Type,View_ImageWidth,View_ImageHeight,Draw_ImageWidth,Draw_ImageHeight,Draw_Graph
	Private Draw_FontColor,Draw_FontFamily,Draw_FontSize,Draw_FontBold,Draw_Info,Draw_Type,Draw_XYType,Draw_SizeType
	Private RName_Str,Transition_Color
	Public ErrCodes,ObjName,UploadFiles,UploadForms,Count,CountSize
	'-----------------------------------------------------------------------------------
	'��ʼ����
	'-----------------------------------------------------------------------------------
	Private Sub Class_Initialize
		SessionName = Empty
		IsBinary = False
		ErrCodes = 0
		Count = 0
		CountSize = 0
		FilePath = "./"
		InceptFile = ""
		FileMaxSize = -1
		MaxFile = 1
		Upload_Type = -1
		Preview_Type = 999
		ObjName = "δ֪���"
		View_ImageWidth = 0
		View_ImageHeight = 0
		Draw_FontColor	= &H000000
		Draw_FontFamily	= "Arial"
		Draw_FontSize	= 10
		Draw_FontBold	= False
		Draw_Info		= "BBS.DVBBS.NET"
		Draw_Type		= -1
		Set UploadFiles = Server.CreateObject ("Scripting.Dictionary")
		Set UploadForms = Server.CreateObject ("Scripting.Dictionary")
		UploadFiles.CompareMode = 1
		UploadForms.CompareMode = 1
	End Sub

	'-----------------------------------------------------------------------------------
	'������
	'-----------------------------------------------------------------------------------
	Private Sub Class_Terminate
		If IsObject(UploadObj) Then
			Set UploadObj = Nothing
		End If
		If IsObject(ImageObj) Then
			Set ImageObj = Nothing
		End If
		UploadFiles.RemoveAll
		UploadForms.RemoveAll
		Set UploadForms = Nothing
		Set UploadFiles = Nothing
	End Sub

	'-----------------------------------------------------------------------------------
	'�����ϴ��Ƿ񷵻��ļ�������
	'-----------------------------------------------------------------------------------
	Public Property Let GetBinary(Byval Values)
		IsBinary = Values
	End Property

	'-----------------------------------------------------------------------------------
	'�����ϴ��������� (�Զ��ŷָ�����ļ�����)
	'-----------------------------------------------------------------------------------
	Public Property Let InceptFileType(Byval Values)
		InceptFile = Lcase(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'�����ϴ��������� (�Զ��ŷָ�����ļ�����)
	'-----------------------------------------------------------------------------------
	Public Property Let ChkSessionName(Byval Values)
		SessionName = Values
	End Property

	'-----------------------------------------------------------------------------------
	'�����ϴ��ļ���С���� (��λ��kb)
	'-----------------------------------------------------------------------------------
	Public Property Let MaxSize(Byval Values)
		FileMaxSize = ChkNumeric(Values) * 1024
	End Property
	Public Property Get MaxSize
		MaxSize = FileMaxSize
	End Property

	'-----------------------------------------------------------------------------------
	'����ÿ���ϴ��ļ�����
	'-----------------------------------------------------------------------------------
	Public Property Let InceptMaxFile(Byval Values)
		MaxFile = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'�����ϴ�Ŀ¼·��
	'-----------------------------------------------------------------------------------
	Public Property Let UploadPath(Byval Path)
		FilePath = Replace(Path,Chr(0),"")
		If Right(FilePath,1)<>"/" Then FilePath = FilePath & "/"
	End Property

	Public Property Get UploadPath
		UploadPath = FilePath
	End Property

	'-----------------------------------------------------------------------------------
	'��ȡ������Ϣ
	'-----------------------------------------------------------------------------------
	Public Property Get Description
		Select Case ErrCodes
			Case 1 : Description = "��֧�� " & ObjName & " �ϴ�������������δ��װ�������"
			Case 2 : Description = "��δѡ���ϴ������"
			Case 3 : Description = "����ѡ����Ҫ�ϴ����ļ�!"
			Case 4 : Description = "�ļ���С���������� " & (FileMaxSize\1024) & "KB!"
			Case 5 : Description = "�ļ����Ͳ���ȷ!"
			Case 6 : Description = "�Ѵﵽ�ϴ��������ޣ�"
			Case 7 : Description = "�벻Ҫ�ظ��ύ��"
			Case Else
				Description = Empty
		End Select
	End Property

	'-----------------------------------------------------------------------------------
	'�����ļ���ǰ׺
	'-----------------------------------------------------------------------------------
	Public Property Let RName(Byval Values)
		RName_Str = Values
	End Property

	'-----------------------------------------------------------------------------------
	'�����ϴ��������
	'-----------------------------------------------------------------------------------
	Public Property Let UploadType(Byval Types)
		Upload_Type = Types
		If Upload_Type = "" or Not IsNumeric(Upload_Type) Then
			Upload_Type = -1
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'�����ϴ�ͼƬ�������
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewType(Byval Types)
		Preview_Type = Types
		On Error Resume Next
		If Preview_Type = "" or Not IsNumeric(Preview_Type) Then
			Preview_Type = 999
		Else
			If PreviewType <> 999 Then
				Select Case Preview_Type
					Case 0
					'---------------------CreatePreviewImage---------------
						ObjName = "CreatePreviewImage���"
						Set ImageObj = Server.CreateObject("CreatePreviewImage.cGvbox")
					Case 1
					'---------------------AspJpegV1.2---------------
						ObjName = "AspJpegV1.2���"
						Set ImageObj = Server.CreateObject("Persits.Jpeg")
					Case 2
					'---------------------SoftArtisans ImgWriter V1.21---------------
						ObjName = "SoftArtisans ImgWriter V1.21���"
						Set ImageObj = Server.CreateObject("SoftArtisans.ImageGen")
					Case Else
						Preview_Type = 999
				End Select
				If Err.Number<>0 Then
					ErrCodes = 1
				End If
			End If
		End If
	End Property

	Public Property Get PreviewType
		PreviewType = Preview_Type
	End Property

	'-----------------------------------------------------------------------------------
	'����Ԥ��ͼƬ��������
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewImageWidth(Byval Values)
		View_ImageWidth = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'����Ԥ��ͼƬ�߶�����
	'-----------------------------------------------------------------------------------
	Public Property Let PreviewImageHeight(Byval Values)
		View_ImageHeight = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡͼƬ�����������������
	'-----------------------------------------------------------------------------------
	Public Property Let DrawImageWidth(Byval Values)
		Draw_ImageWidth = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡͼƬ����������߶�����
	'-----------------------------------------------------------------------------------
	Public Property Let DrawImageHeight(Byval Values)
		Draw_ImageHeight = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡͼƬ����������͸��������
	'-----------------------------------------------------------------------------------
	Public Property Let DrawGraph(Byval Values)
		If IsNumeric(Values) Then
			Draw_Graph = Formatnumber(Values,2)
		Else
			Draw_Graph = 1
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡͼƬ͸����ȥ����ɫֵ
	'-----------------------------------------------------------------------------------
	Public Property Let TransitionColor(Byval Values)
		If Values<>"" or Values<>"0" Then
			Transition_Color = Replace(Values,"#","&h")
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡ������ɫ
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontColor(Byval Values)
		If Values<>"" or Values<>"0" Then
			Draw_FontColor = Replace(Values,"#","&h")
		End If
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡ���������ʽ
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontFamily(Byval Values)
		Draw_FontFamily = Values
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡ���������С
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontSize(Byval Values)
		Draw_FontSize = Values
	End Property

	'-----------------------------------------------------------------------------------
	'����ˮӡ�����Ƿ���� Boolean
	'-----------------------------------------------------------------------------------
	Public Property Let DrawFontBold(Byval Values)
		Draw_FontBold = ChkBoolean(Values)
	End Property
	'-----------------------------------------------------------------------------------
	'����ˮӡ������Ϣ��ͼƬ��Ϣ
	'-----------------------------------------------------------------------------------
	Public Property Let DrawInfo(Byval Values)
		Draw_Info = Values
	End Property

	'-----------------------------------------------------------------------------------
	'����ģʽ��0=������ˮӡ ��1=����ˮӡ���� ��2=����ˮӡͼƬ
	'-----------------------------------------------------------------------------------
	Public Property Let DrawType(Byval Values)
		Draw_Type = ChkNumeric(Values)
	End Property

	'-----------------------------------------------------------------------------------
	'ͼƬ����ˮӡLOGOλ�����꣺"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
	'-----------------------------------------------------------------------------------
	Public Property Let DrawXYType(Byval Values)
		 Draw_XYType = Values
	End Property

	'-----------------------------------------------------------------------------------
	'����Ԥ��ͼƬ��С����"0"=�̶���С��"1"=�ȱ�����С
	'-----------------------------------------------------------------------------------
	Public Property Let DrawSizeType(Byval Values)
		Draw_SizeType = Values
	End Property

	Private Function ChkNumeric(Byval Values)
		If Values<>"" and Isnumeric(Values) Then
			ChkNumeric = Int(Values)
		Else
			ChkNumeric = 0
		End If
	End Function

	Private Function ChkBoolean(Byval Values)
		If Typename(Values)="Boolean" or IsNumeric(Values) or Lcase(Values)="false" or Lcase(Values)="true" Then
			ChkBoolean = CBool(Values)
		Else
			ChkBoolean = False
		End If
	End Function

	'-----------------------------------------------------------------------------------
	'����ʱ�䶨���ļ���
	'-----------------------------------------------------------------------------------
	Private Function FormatName(Byval FileExt)
		Dim RanNum,TempStr
		Randomize
		RanNum = Int(90000*rnd)+10000
		TempStr = Year(now) & Month(now) & Day(now) & Hour(now) & Minute(now) & Second(now) & RanNum & "." & FileExt
		If RName_Str<>"" Then
			TempStr = RName_Str & TempStr
		End If
		FormatName = TempStr
	End Function
	
	'-----------------------------------------------------------------------------------
	'��ʽ��׺
	'-----------------------------------------------------------------------------------
	Private Function FixName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		FixName = Lcase(UpFileExt)
		FixName = Replace(FixName,Chr(0),"")
		FixName = Replace(FixName,".","")
		FixName = Replace(FixName,"'","")
		FixName = Replace(FixName,"asp","")
		FixName = Replace(FixName,"asa","")
		FixName = Replace(FixName,"aspx","")
		FixName = Replace(FixName,"cer","")
		FixName = Replace(FixName,"cdx","")
		FixName = Replace(FixName,"htr","")
	End Function

	'-----------------------------------------------------------------------------------
	'�ж��ļ������Ƿ�ϸ�
	'-----------------------------------------------------------------------------------
	Private Function CheckFileExt(FileExt)
		Dim Forumupload,i
		CheckFileExt=False
		If FileExt="" or IsEmpty(FileExt) Then
			CheckFileExt = False
			Exit Function
		End If
		If FileExt="asp" or FileExt="asa" or FileExt="aspx" Then
			CheckFileExt = False
			Exit Function
		End If
		Forumupload = Split(InceptFile,",")
		For i = 0 To ubound(Forumupload)
			If FileExt = Trim(Forumupload(i)) Then
				CheckFileExt = True
				Exit Function
			Else
				CheckFileExt = False
			End If
		Next
	End Function

	'-----------------------------------------------------------------------------------
	'�ж��ļ�����:0=����,1=ͼƬ,2=FLASH,3=����,4=��Ӱ
	'-----------------------------------------------------------------------------------
	Private Function CheckFiletype(Byval FileExt)
		FileExt = Lcase(Replace(FileExt,".",""))
		Select Case FileExt
				Case "gif", "jpg", "jpeg","png","bmp","tif","iff"
					CheckFiletype=1
				Case "swf", "swi"
					CheckFiletype=2
				Case "mid", "wav", "mp3","rmi","cda"
					CheckFiletype=3
				Case "avi", "mpg", "mpeg","ra","ram","wov","asf"
					CheckFiletype=4
				Case Else
					CheckFiletype=0
		End Select
	End Function

	'-----------------------------------------------------------------------------------
	'ִ�б����ϴ��ļ�
	'-----------------------------------------------------------------------------------
	Public Sub SaveUpFile()
		On Error Resume Next
		Select Case (Upload_Type) 
			Case 0
				ObjName = "�����"
				Set UploadObj = New UpFile_Class
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_0
				End If
			Case 1
				ObjName = "Aspupload3.0���"
				Set UploadObj = Server.CreateObject("Persits.Upload") 
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_1
				End If
			Case 2
				ObjName = "SA-FileUp 4.0���"
				Set UploadObj = Server.CreateObject("SoftArtisans.FileUp")
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_2
				End If
			Case 3
				ObjName = "DvFile.Upload V1.0���"
				Set UploadObj = Server.CreateObject("DvFile.Upload")
				If Err.Number<>0 Then
					ErrCodes = 1
				Else
					SaveFile_3
				End If
			Case Else
				ErrCodes = 2
		End Select
	End Sub

	''-----------------------------------------------------------------------------------
	' �ϴ���������
	''-----------------------------------------------------------------------------------
	''-----------------------------------------------------------------------------------
	''������ϴ�
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_0()
		Dim FormName,Item,File
		Dim FileExt,FileName,FileType,FileToBinary
		UploadObj.InceptFileType = InceptFile
		UploadObj.MaxSize = FileMaxSize
		UploadObj.GetDate ()	'ȡ���ϴ�����
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		If UploadObj.Err > 0 then
			Select Case UploadObj.Err
				Case 1 : ErrCodes = 3
				Case 2 : ErrCodes = 4
				Case 3 : ErrCodes = 5
			End Select
			Exit Sub
		Else
			For Each FormName In UploadObj.File		''�г������ϴ��˵��ļ�
				If Count>MaxFile Then
					ErrCodes = 6
					Exit Sub
				End If
				Set File = UploadObj.File(FormName)
				FileExt = FixName(File.FileExt)
				If CheckFileExt(FileExt) = False then
					ErrCodes = 5
					EXIT SUB
				End If
				FileName = FormatName(FileExt)
				FileType = CheckFiletype(FileExt)
				If IsBinary Then
					FileToBinary = File.FileData
				End If
				If File.FileSize>0 Then
					File.SaveToFile Server.Mappath(FilePath & FileName)
					AddData FormName , _ 
							FileName , _
							FilePath , _
							File.FileSize , _
							File.FileType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.FileWidth , _
							File.FileHeight
					Count = Count + 1
					CountSize = CountSize + File.FileSize
				End If
				Set File=Nothing
			Next
			For Each Item in UploadObj.Form
				If UploadForms.Exists (Item) Then _
					UploadForms(Item) = UploadForms(Item) & ", " & UploadObj.Form(Item) _
				Else _
				UploadForms.Add Item , UploadObj.Form(Item)
			Next
			If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub
	''-----------------------------------------------------------------------------------
	''Aspupload3.0����ϴ�
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_1()
		Dim FileCount
		Dim FormName,Item,File
		Dim FileExt,FileName,FileType,FileToBinary
		UploadObj.OverwriteFiles = False		'���ܸ���
		UploadObj.IgnoreNoPost = True
		UploadObj.SetMaxSize FileMaxSize, True	'���ƴ�С
		FileCount = UploadObj.Save
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If

		If Err.Number = 8 Then
				ErrCodes = 4
				EXIT SUB
		Else 
				If Err <> 0 Then
					ErrCodes = -1
					Response.Write "������Ϣ: " & Err.Description
					EXIT SUB
				End If
				If FileCount < 1 Then 
					ErrCodes = 3
					EXIT SUB
				End If
				For Each File In UploadObj.Files	'�г������ϴ��ļ�
					If Count>MaxFile Then
						ErrCodes = 6
						Exit Sub
					End If
					FileExt = FixName(Replace(File.Ext,".",""))
					If CheckFileExt(FileExt) = False then
						ErrCodes = 5
						EXIT SUB
					End If
					FileName = FormatName(FileExt)
					FileType = CheckFiletype(FileExt)
					If IsBinary Then
						FileToBinary = File.Binary
					End If
					'File.Filename
					If File.Size>0 Then
						File.SaveAs Server.Mappath(FilePath & FileName)
						AddData File.Name , _ 
							FileName , _
							FilePath , _
							File.Size , _
							File.ContentType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.ImageWidth , _
							File.ImageHeight
						Count = Count + 1
						CountSize = CountSize + File.Size
					End If
				Next
				For Each Item in UploadObj.Form
					If UploadForms.Exists (Item) Then _
						UploadForms(Item) = UploadForms(Item) & ", " & Item.Value _
					Else _
						UploadForms.Add Item.Name , Item.Value
				Next
				If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub
	''-----------------------------------------------------------------------------------
	''SA-FileUp 4.0����ϴ�FileUpSE V4.09
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_2()
		Dim FormName,Item,File,FormNames
		Dim FileExt,FileName,FileType,FileToBinary
		Dim Filesize
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		For Each FormName In UploadObj.Form
			FormNames = ""
			If IsObject(UploadObj.Form(FormName)) Then
				If Not UploadObj.Form(FormName).IsEmpty Then
					UploadObj.Form(FormName).Maxbytes = FileMaxSize	'���ƴ�С
					UploadObj.OverWriteFiles = False
					Filesize = UploadObj.Form(FormName).TotalBytes
					If Err.Number<>0 Then
						ErrCodes = -1
						Response.Write "������Ϣ: " & Err.Description
						EXIT SUB
					End If
					If Filesize>FileMaxSize then
						ErrCodes = 4
						Exit sub
					End If
					FileName	= UploadObj.Form(FormName).ShortFileName	 'ԭ�ļ���
					FileExt		= Mid(Filename, InStrRev(Filename, ".")+1)
					FileExt		= FixName(FileExt)
					If CheckFileExt(FileExt) = False then
						ErrCodes = 5
						EXIT SUB
					End If
					FileName = FormatName(FileExt)
					FileType = CheckFiletype(FileExt)
					'If IsBinary Then
						'FileToBinary = UploadContents (2)
					'End If
					'�����ļ�
					If Filesize>0 Then
						UploadObj.Form(FormName).SaveAs Server.MapPath(FilePath & FileName)
						AddData FormName , _ 
								FileName , _
								FilePath , _
								FileSize , _
								UploadObj.Form(FormName).ContentType , _
								FileType , _
								FileToBinary , _
								FileExt , _
								-1 , _
								-1
						Count = Count + 1
						CountSize = CountSize + Filesize
					End If
				Else
					ErrCodes = 3
					EXIT SUB
				End If
			Else
				If UploadObj.FormEx(FormName).Count > 1 Then
					For Each FormNames In UploadObj.FormEx(FormName)
						FormNames = FormNames & ", " & FormNames
					Next
					UploadForms.Add FormName , FormNames
				Else
					UploadForms.Add FormName , UploadObj.Form(FormName)
				End If
			End If
		Next
		If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
	End Sub
	''-----------------------------------------------------------------------------------
	''DvFile.Upload V1.0����ϴ�
	''-----------------------------------------------------------------------------------
	Private Sub SaveFile_3()
		Dim FormName,Item,File
		Dim FileExt,FileName,FileType,FileToBinary
		UploadObj.InceptFileType = InceptFile
		UploadObj.MaxSize = FileMaxSize
		UploadObj.Install
		FileToBinary = Null
		If Not IsEmpty(SessionName) Then
			If Session(SessionName) <> UploadObj.Form(SessionName) or Session(SessionName) = Empty Then
				ErrCodes = 7
				Exit Sub
			End If
		End If
		If UploadObj.Err > 0 then
			Select Case UploadObj.Err
				Case 1 : ErrCodes = 3
				Case 2 : ErrCodes = 4
				Case 3 : ErrCodes = 5
				Case 4 : ErrCodes = 5
				Case 5 : ErrCodes = -1
			End Select
			Exit Sub
		Else
			For Each FormName In UploadObj.File		''�г������ϴ��˵��ļ�
				If Count>MaxFile Then
					ErrCodes = 6
					Exit Sub
				End If
				Set File = UploadObj.File(FormName)
				FileExt = FixName(File.FileExt)
				If CheckFileExt(FileExt) = False then
					ErrCodes = 5
					EXIT SUB
				End If
				FileName = FormatName(FileExt)
				FileType = CheckFiletype(FileExt)
				If IsBinary Then
					FileToBinary = File.FileData
				End If
				If File.FileSize>0 Then
					UploadObj.SaveToFile Server.mappath(FilePath & FileName),FormName
					AddData FormName , _ 
							FileName , _
							FilePath , _
							File.FileSize , _
							File.FileType , _
							FileType , _
							FileToBinary , _
							FileExt , _
							File.FileWidth , _
							File.FileHeight
					Count = Count + 1
					CountSize = CountSize + File.FileSize
				End If
				Set File=Nothing
			Next
			For Each Item in UploadObj.Form
				UploadForms.Add Item.Name , Item.Value
			Next
			If Not IsEmpty(SessionName) Then Session(SessionName) = Empty
		End If
	End Sub

	Private Sub AddData( Form_Name,File_Name,File_Path,File_Size,File_ContentType,File_Type,File_Data,File_Ext,File_Width,File_Height )
		Set FileInfo = New FileInfo_Cls
			FileInfo.FormName = Form_Name
			FileInfo.FileName = File_Name
			FileInfo.FilePath = File_Path
			FileInfo.FileSize = File_Size
			FileInfo.FileType = File_Type
			FileInfo.FileContentType = File_ContentType
			FileInfo.FileExt = File_Ext
			FileInfo.FileData = File_Data
			FileInfo.FileHeight = File_Height
			FileInfo.FileWidth = File_Width
			UploadFiles.Add Form_Name , FileInfo
		Set FileInfo = Nothing
	End Sub

	'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
	Public Sub CreateView(Imagename,TempFilename,FileExt)
		If ErrCodes <>0 Then Exit Sub
		Select Case Preview_Type
			Case 0
				Image_Obj_0 Imagename,TempFilename,FileExt
			Case 1
				Image_Obj_1 Imagename,TempFilename,FileExt
			Case 2
				Image_Obj_2 Imagename,TempFilename,FileExt
			Case Else
				Preview_Type = 999
		End Select
	End Sub

	Sub Image_Obj_0(Imagename,TempFilename,FileExt)
			ImageObj.SetSavePreviewImagePath = Server.MapPath(TempFilename)			'Ԥ��ͼ���·��
			ImageObj.SetPreviewImageSize = SetPreviewImageSize						'Ԥ��ͼ����
			ImageObj.SetImageFile = Trim(Server.MapPath(Imagename))					'Imagenameԭʼ�ļ�������·��
			'����Ԥ��ͼ���ļ�
			If ImageObj.DoImageProcess = False Then
				ErrCodes = -1
				Response.Write "����Ԥ��ͼ����: " & ImageObj.GetErrString
			End If
	End Sub

	'---------------------AspJpegV1.2---------------
	Sub Image_Obj_1(Imagename,TempFilename,FileExt)
			' ��ȡҪ������ԭ�ļ�
			Dim Draw_X,Draw_Y,Logobox
			Draw_X = 0
			Draw_Y = 0
			FileExt = Lcase(FileExt)
			On Error Resume Next
			ImageObj.Open Trim(Server.MapPath(Imagename))
			If Err Then
				err.Clear
				Exit Sub
			End If
			If ImageObj.OriginalWidth<View_ImageWidth or ImageObj.Originalheight<View_ImageHeight Then
				TempFilename = ""
				Exit Sub
			Else
				If FileExt<>"gif" and ImageObj.OriginalWidth > Draw_ImageWidth * 2 and Draw_Type >0 Then
					Draw_X = DrawImage_X(ImageObj.OriginalWidth,Draw_ImageWidth,2)
					Draw_Y = DrawImage_y(ImageObj.Originalheight,Draw_ImageHeight,2)
					If Draw_Type=2 Then
						Set Logobox = Server.CreateObject("Persits.Jpeg")
						'*����ˮӡͼƬ	����ʱ��ر�ˮӡ����*
						'//��ȡ���ӵ�ͼƬ
						Logobox.Open Server.MapPath(Draw_Info)
						Logobox.Width = Draw_ImageWidth								'// ����ͼƬ��ԭ����
						Logobox.Height = Draw_ImageHeight							'// ����ͼƬ��ԭ�߶�
						ImageObj.DrawImage Draw_X, Draw_Y, Logobox, Draw_Graph,Transition_Color,90	'// ����ͼƬ��λ�ü����꣨����ˮӡͼƬ��
						'ImageObj.Sharpen 1, 130
						ImageObj.Save Server.MapPath(Imagename)
						Set Logobox=Nothing
					Else
						'//�����޸����弰������ɫ��
						ImageObj.Canvas.Font.Color		= Draw_FontColor	'// ���ֵ���ɫ
						ImageObj.Canvas.Font.Family		= Draw_FontFamily	'// ���ֵ�����
						ImageObj.Canvas.Font.Bold		= Draw_FontBold
						ImageObj.Canvas.Font.Size		= Draw_FontSize					'//�����С
						' Draw frame: black, 2-pixel width
						ImageObj.Canvas.Print Draw_X, Draw_Y, Draw_Info	'// �������ֵ�λ������
						ImageObj.Canvas.Pen.Color		= &H000000		'// �߿����ɫ
						ImageObj.Canvas.Pen.Width		= 1				'// �߿�Ĵ�ϸ
						ImageObj.Canvas.Brush.Solid	= False			'// ͼƬ�߿����Ƿ������ɫ
						'ImageObj.Canvas.Bar 0, 0, ImageObj.Width, ImageObj.Height	'// ͼƬ�߿��ߵ�λ������
						ImageObj.Save Server.MapPath(Imagename)
					End If
				End If
				If ImageObj.Width > ImageObj.height Then
					ImageObj.Width = View_ImageWidth
					ImageObj.Height = ViewImage_Height(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
				Else
					ImageObj.Width = ViewImage_Width(ImageObj.OriginalWidth,ImageObj.Originalheight,View_ImageWidth,View_ImageHeight)
					ImageObj.Height = View_ImageHeight
				End If
				ImageObj.Sharpen 1, 120
				ImageObj.Save Server.MapPath(TempFilename)		'// ����Ԥ���ļ�
			End If
	End Sub

	'SoftArtisans ImgWriter V1.21
	Public Sub Image_Obj_2(Imagename,TempFilename,FileExt)
			'�������
			Dim Draw_X,Draw_Y
			FileExt = Lcase(FileExt)
			Draw_X = 0
			Draw_Y = 0
			' ��ȡҪ������ԭ�ļ�
			ImageObj.LoadImage Trim(Server.MapPath(Imagename))
			If ImageObj.ErrorDescription <> "" Then
				TempFilename = ""
				ErrCodes = -1
				Response.Write "����Ԥ��ͼ����: " &ImageObj.ErrorDescription
				Exit Sub
			End If
			If ImageObj.Width<Cint(View_ImageWidth) or ImageObj.Height<Cint(View_ImageHeight) Then
				TempFilename=""
				Exit Sub
			Else
				IF FileExt<>"gif" and ImageObj.Width > Draw_ImageWidth * 2 and Draw_Type>0 Then
					Draw_X = DrawImage_X(ImageObj.Width,Draw_ImageWidth,2)
					Draw_Y = DrawImage_y(ImageObj.Height,Draw_ImageHeight,2)
					Dim saiTopMiddle
					Select Case Draw_XYType
						Case "0" '����
							saiTopMiddle = 3
						Case "1" '����
							saiTopMiddle = 5
						Case "2" '����
							saiTopMiddle = 1
						Case "3" '����
							saiTopMiddle = 6
						Case "4" '����
							saiTopMiddle = 8
						Case Else '����ʾ
							saiTopMiddle = 0
					End Select
					If Draw_Type=2 Then
						ImageObj.AddWatermark Server.MapPath(Draw_Info), saiTopMiddle, Draw_Graph,Transition_Color,True
						'ImageObj.AddWatermark Server.MapPath(Request.QueryString("mimg")), 0, 0.3
					Else
						ImageObj.Font.Italic	= False			'б��
						ImageObj.Font.height	= Draw_FontSize
						ImageObj.Font.name		= Draw_FontFamily
						ImageObj.Font.Color		= Draw_FontColor
						ImageObj.Text			= Draw_Info
						ImageObj.DrawTextOnImage Draw_X, Draw_Y, ImageObj.TextWidth, ImageObj.TextHeight
					End If
					ImageObj.SaveImage 0, ImageObj.ImageFormat, Server.MapPath(Imagename)
				End If
				'ImageObj.SharpenImage 100
				ImageObj.ColorResolution = 24	'24ɫ����
				ImageObj.ResizeImage View_ImageWidth,View_ImageHeight,0,0
				'0=saiFile,1=saiMemory,2=saiBrowser,4=saiDatabaseBlob
				'saiBMP=1,saiGIF=2,saiJPG=3,saiPNG=4,saiPCX=5,saiTIFF=6,saiWMF=7,saiEMF=8,saiPSD=9 
				ImageObj.SaveImage 0, 3, Server.MapPath(TempFilename)
			End If
	End Sub

	'������̶���С
	Private Function ViewImage_Width(Image_W,Image_H,xView_W,xView_H)
		If Draw_SizeType = "1" Then
			ViewImage_Width = Image_W * xView_H / Image_H
		Else
			ViewImage_Width = xView_W
		End If
	End Function

	Private Function ViewImage_Height(Image_W,Image_H,xView_W,xView_H)
		If Draw_SizeType = "1" Then
			ViewImage_Height = xView_W * Image_H / Image_W
		Else
			ViewImage_Height = xView_H
		End If
	End Function

	'SpaceVal X�������Ե����
	Private Function DrawImage_X(xImage_W,xLogo_W,SpaceVal)
		Select Case Draw_XYType
			Case "0" '����
				DrawImage_X = SpaceVal
			Case "1" '����
				DrawImage_X = SpaceVal
			Case "2" '����
				DrawImage_X = (xImage_W + xLogo_W) / 2
			Case "3" '����
				DrawImage_X = xImage_W - xLogo_W - SpaceVal
			Case "4" '����
				DrawImage_X = xImage_W - xLogo_W - SpaceVal
			Case Else '����ʾ
				DrawImage_X = 0
		End Select
	End Function

	'SpaceVal Y�������Ե����
	Private Function DrawImage_Y(yImage_H,yLogo_H,SpaceVal)
		Select Case Draw_XYType
			Case "0" '����
				DrawImage_Y = SpaceVal
			Case "1" '����
				DrawImage_Y = yImage_H - yLogo_H - SpaceVal
			Case "2" '����
				DrawImage_Y = (yImage_H + yLogo_H) / 2
			Case "3" '����
				DrawImage_Y = SpaceVal
			Case "4" '����
				DrawImage_Y = yImage_H - yLogo_H - SpaceVal
			Case Else '����ʾ
				DrawImage_Y = 0
		End Select
	End Function

End Class

Class FileInfo_Cls
	Public FormName,FileName,FilePath,FileSize,FileContentType,FileType,FileData,FileExt,FileWidth,FileHeight
	Private Sub Class_Initialize
		FileWidth = -1
		FileHeight = -1
	End Sub
End Class
%>