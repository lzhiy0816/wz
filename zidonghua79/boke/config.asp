<%
'����ϵͳ�Ƿ�������ݿ�,1��ʾ�������ݿ�,0��ʾ�ڶ�����̳���ݿ���
Const Dv_Boke_InDvbbsData = 1
'����ϵͳ���ݿ�����,0��ʾACCESS,1��ʾSQL Server
Const Dv_Boke_DataBase = 0
'�Ƿ�֧��ģ��HTML,��Ҫ������֧��Isapiwrite�����������
Const Is_Isapi_Rewrite = 0

Dim bSqlNowString,DvBoke,Boke_Conn
If Dv_Boke_InDvbbsData = 1 Then
	If Dv_Boke_DataBase = 1 Then
		bSqlNowString = "GetDate()"
	Else
		bSqlNowString = "Now()"
	End If
Else
	bSqlNowString = SqlNowString
End If
MyDbPath = ""

'��������ϵͳ���ݿ����Ӻ���
Sub Boke_ConnectionDatabase()
	Dim ConnStr
	If Dv_Boke_InDvbbsData = 0 Then
		If Not IsObject(Conn) Then ConnectionDatabase
		Set Boke_Conn = Conn
		Exit Sub
	End If
	If Dv_Boke_DataBase = 1 Then
		'sql���ݿ����Ӳ��������ݿ������û����롢�û�������������������local�������IP��
		Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
		SqlDatabaseName = "dvbbs7"
		SqlPassword = "dvbbs"
		SqlUsername = "dvbbs"
		SqlLocalName = "(local)"
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		Dim Db
		'����û���һ��ʹ�����޸ı������ݿ��ַ����Ӧ�޸�dataĿ¼�����ݿ�����
		Db = MyDbPath & "Boke/data/Dvboke.mdb"
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
	End If
	'On Error Resume Next
	Set Boke_Conn = Server.CreateObject("ADODB.Connection")
	Boke_Conn.Open ConnStr
	If Err Then
		err.Clear
		Set Boke_Conn = Nothing
		Response.Write "������ݿ����ӳ������������ִ���"'ע�ͣ���Ҫ���⼸���ַ����Ӣ�ġ�
		Response.End
	End If
End Sub
%>
<!--#Include File="Cls_Main.asp"-->
<%
Set DvBoke = New Cls_DvBoke
%>