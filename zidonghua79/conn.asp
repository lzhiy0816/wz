<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Response.Buffer = True
Dim Startime
Dim SqlNowString,Dvbbs,template,MyBoardOnline
Dim Conn,Plus_Conn,Db,MyDbPath
Const fversion="7.1.0 Sp1"
Const EnabledSession= True
Startime = Timer()
'ϵͳ����XML�汾����
'��߰汾Ϊ.4.0 ����Ϊ�� Const MsxmlVersion=".3.0" Const MsxmlVersion=".2.6" ��Ͱ汾Const MsxmlVersion=""
Const MsxmlVersion=".3.0"
'���޸�����һ��========================�������ݿ����1ΪSQL���ݿ⣬0ΪAccess���ݿ�=============================
Const IsSqlDataBase = 0
MyDbPath = ""
'================================================================================================================
If IsSqlDataBase = 1 Then
'���޸����ö���========================SQL���ݿ�����=============================================================
'sql���ݿ����Ӳ��������ݿ���(SqlDatabaseName)���û�����(SqlPassword)���û���(SqlUsername)��
'������(SqlLocalName)��������local�������IP��
Const SqlDatabaseName = "dvbbs"
Const SqlPassword = "dvbbs"
Const SqlUsername = "dvbbs"
Const SqlLocalName = "(local)"
'================================================================================================================
SqlNowString = "GetDate()"
Else
'���޸���������========================Access���ݿ�����==========================================================
'����û���һ��ʹ�����޸ı������ݿ��ַ����Ӧ�޸�dataĿ¼�����ݿ����ƣ���:��dvbbs6.mdb�޸�Ϊdvbbs6.asp
Db = "data/dianqizidonghua79.mdb"
'================================================================================================================
SqlNowString = "Now()"
End If
Const IsDeBug = 1
Set Dvbbs = New Cls_Forum
Set template = New cls_templates
Sub ConnectionDatabase
	Dim ConnStr
	If IsSqlDataBase = 1 Then
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & db)
	End If
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "���ݿ����ӳ��������������ִ���"'ע�ͣ���Ҫ���⼸���ַ����Ӣ�ġ�
		Response.End
	End If
End Sub
'-----------------------------------------------------------------------------------------------------
'�������߿���������
Sub Plus_ConnectionDatabase
	Dim ConnStr
	If IsSqlDataBase = 1 Then
		'sql���ݿ����Ӳ��������ݿ������û����롢�û�������������������local�������IP��
		Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
		SqlDatabaseName = "dvbbs7"
		SqlPassword = "dvbbs"
		SqlUsername = "dvbbs"
		SqlLocalName = "(local)"
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		Dim Db
		'����û���һ��ʹ�����޸ı������ݿ��ַ����Ӧ�޸�dataĿ¼�����ݿ����ƣ��罫dvbbs6.mdb�޸�Ϊdvbbs6.asp
		Db = MyDbPath & "data/Dv_Plus_Tools.mdb"
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
	End If
	On Error Resume Next
	Set Plus_Conn = Server.CreateObject("ADODB.Connection")
	Plus_Conn.open ConnStr
	If Err Then
		err.Clear
		Set Plus_Conn = Nothing
		Response.Write "������ݿ����ӳ��������������ִ���"'ע�ͣ���Ҫ���⼸���ַ����Ӣ�ġ�
		Response.End
	End If
End Sub
'-----------------------------------------------------------------------------------------------------
%>