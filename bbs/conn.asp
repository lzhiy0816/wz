<%@ LANGUAGE=VBScript CodePage=936%>
<%Response.Buffer=True%>
<%
IsSqlDataBase=0  '�������ݿ����0ΪAccess���ݿ⣬1ΪSQL���ݿ�
If IsSqlDataBase=0 Then
'''''''''''''''''''''''''''''' Access���ݿ� '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
datapath    ="lzydata/"      '���ݿ�Ŀ¼�����·��
datafile    ="lzybbs.mdb"      '���ݿ���ļ���
Connstr="DBQ="&server.mappath(""&datapath&""&datafile&"")&";DRIVER={Microsoft Access Driver (*.mdb)};"
SqlNowString="Now()"
SqlChar="'"
ver="4.00"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Else
'''''''''''''''''''''''''''''' SQL���ݿ� ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SqlLocalName   ="(local)"     '����IP  [ ������ (local) �����IP ]
SqlUsername    ="sa"          '�û���
SqlPassword    ="1"           '�û�����
SqlDatabaseName="bbsxp"       '���ݿ���
ConnStr = "Provider=Sqloledb; User ID=" & SqlUsername & "; Password=" & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source=" & SqlLocalName & ";"
SqlNowString="GetDate()"
ver="4.00 SQL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End If
%>