<%
'����Ȧ��ϵͳ�Ƿ�������ݿ�,1��ʾ�������ݿ�,0��ʾ�ڶ�����̳���ݿ���
Const Dv_IndivGroup_DataBaseFlag = 0
'����Ȧ��ϵͳ���ݿ�����,0��ʾACCESS,1��ʾSQL Server
Const Dv_IndivGroup_DataBaseType = 0

Dim Dv_IndivGroup_Class,Dv_IndivGroup_Conn

'����Ȧ��ϵͳ���ݿ����Ӻ���
Sub Dv_IndivGroup_ConnectionDatabase()
	Dim ConnStr,DB
	Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
	If Dv_IndivGroup_DataBaseFlag = 0 Then
		If Not IsObject(Conn) Then ConnectionDatabase
		Set Dv_IndivGroup_Conn = Conn
	Else
		If Dv_IndivGroup_DataBaseFlag = 1 Then
			'SQL���ݿ����Ӳ��������ݿ������û����롢�û�������������������(local)�������IP��
			SqlDatabaseName = "dvbbs8"
			SqlPassword = "dvbbs"
			SqlUsername = "dvbbs"
			SqlLocalName = "(local)"
			ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
		Else
			'����û���һ��ʹ�����޸ı������ݿ��ַ����Ӧ�޸�dataĿ¼�����ݿ�����
			DB = "IndivGroup/Data/Dv_IndivGroup.mdb"
			ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(DB)
		End If
		'On Error Resume Next
		Set Dv_IndivGroup_Conn = Dvbbs.iCreateObject("ADODB.Connection")
		Dv_IndivGroup_Conn.Open ConnStr
		If Err Then
			err.Clear
			Set Dv_IndivGroup_Conn = Nothing
			Response.Write "������ݿ����ӳ��������������ִ���"'ע�ͣ���Ҫ���⼸���ַ����Ӣ�ġ�
			Response.End
		End If
	End If
End Sub
%>