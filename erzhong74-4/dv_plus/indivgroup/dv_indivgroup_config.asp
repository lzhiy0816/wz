<%
'个性圈子系统是否独立数据库,1表示独立数据库,0表示在动网论坛数据库中
Const Dv_IndivGroup_DataBaseFlag = 0
'个性圈子系统数据库类型,0表示ACCESS,1表示SQL Server
Const Dv_IndivGroup_DataBaseType = 0

Dim Dv_IndivGroup_Class,Dv_IndivGroup_Conn

'个性圈子系统数据库链接函数
Sub Dv_IndivGroup_ConnectionDatabase()
	Dim ConnStr,DB
	Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
	If Dv_IndivGroup_DataBaseFlag = 0 Then
		If Not IsObject(Conn) Then ConnectionDatabase
		Set Dv_IndivGroup_Conn = Conn
	Else
		If Dv_IndivGroup_DataBaseFlag = 1 Then
			'SQL数据库连接参数：数据库名、用户密码、用户名、连接名（本地用(local)，外地用IP）
			SqlDatabaseName = "dvbbs8"
			SqlPassword = "dvbbs"
			SqlUsername = "dvbbs"
			SqlLocalName = "(local)"
			ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
		Else
			'免费用户第一次使用请修改本处数据库地址并相应修改data目录中数据库名称
			DB = "IndivGroup/Data/Dv_IndivGroup.mdb"
			ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(DB)
		End If
		'On Error Resume Next
		Set Dv_IndivGroup_Conn = Dvbbs.iCreateObject("ADODB.Connection")
		Dv_IndivGroup_Conn.Open ConnStr
		If Err Then
			err.Clear
			Set Dv_IndivGroup_Conn = Nothing
			Response.Write "插件数据库连接出错，请检查连接字串。"'注释，需要把这几个字翻译成英文。
			Response.End
		End If
	End If
End Sub
%>