<!--#include file = "../conn.asp"-->
<!-- #include file = "inc/const.asp" -->
<%
Head()
dim admin_flag
dim objFSO
dim uploadfolder
dim uploadfiles
dim upname
dim uid,faceid
dim usernames
dim userface,dnum
dim upfilename
dim pagesize, page,filenum, pagenum
admin_flag = ",34,"
If not Dvbbs.master or instr(","&session("flag")&",",admin_flag) = 0 Then
	Errmsg = ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href = ../admin_login.asp target = _top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
	dvbbs_error()
else
	call main()
	If Errmsg<>"" Then Dvbbs_Error()
	Footer()
End If

sub main()
%>
<table width = "95%" border = "0" cellspacing = "1" cellpadding = "3"  align = center>
<tr>
<td valign = top>
ע�⣺��������Ҫ��������FSOȨ�ޣ�FSO��ذ����뿴΢�������ĵ�<BR>
�����������Թ�����̳�����û��Զ���ͷ���ϴ��ļ��������û�ͷ�������û�ID��������<BR>
�û�ID�Ļ�ÿ���ͨ���û���Ϣ��������������û���Ȼ������Ƶ��û��������ϣ��鿴�������ԣ�����UserID = ��������û���ID
</td>
</tr>
</table>
<table width = "95%" border = "0" cellspacing = "1" cellpadding = "3"  align = center class = "tableBorder" style = "table-layout:fixed;word-break:break-all">
<tr align = center><th width = "*" height = 25>�ļ���</th><th width = "100">�����û�</th><th width = "50">��С</th><th width = "120">������</th><th width = "120">�ϴ�����</th><th width = "35">����</th></tr>
<form method="POST" action="?action=delall">
<%
pagesize = 20
page = request.querystring("page")
If page = "" or not isnumeric(page) Then
	page = 1
Else
	page = int(page)
End If

If trim(request("action"))<>"" Then
	If trim(request("action")) = "delall" Then
		call delface()
	Else
		call maininfo()
	End If
Else
	call maininfo()
End If
call foot()
End Sub

sub maininfo()
On Error Resume Next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
If Err Then
	ErrMsg = "<li>����ϵͳ��֧��FSO�ļ���д������ʹ�ô˹��ܡ�"
	Exit Sub
End If
If request("filename")<>"" Then
	If UpUserFaceFolder="" Then 
		objFSO.DeleteFile(Server.MapPath("../uploadFace/"&request("filename")))
	Else
		objFSO.DeleteFile(Server.MapPath(".."&UpUserFaceFolder&request("filename")))
	End If
End If
If UpUserFaceFolder="" Then 
	Set uploadFolder = objFSO.GetFolder(Server.MapPath("../uploadFace/"))
Else
	Set uploadFolder = objFSO.GetFolder(Server.MapPath(".."&UpUserFaceFolder))
End If
If Err Then
	ErrMsg = "<li>��ʹ�õ��ϴ�ͷ��Ŀ¼����ϵͳĬ��Ŀ¼�����ܽ��й�����"
	Exit Sub
End If
Set uploadFiles = uploadFolder.Files
filenum = uploadfiles.count
pagenum = int(filenum/pagesize)
If filenum mod pagesize>0 Then
	pagenum = pagenum+1
End If
If page> pagenum Then
	page = 1
End If
i = 0
For Each Upname In uploadFiles
	i = i+1
	If i>(page-1)*pagesize and i <= page*pagesize Then
	upfilename = "../uploadFace/"&upname.name
		If instr(upname.name,"_") Then    'ȡ��ͷ����û���
			uid = split(upname.name,"_")
			faceid = uid(0)
			If  IsNumeric(faceid)	then
				set rs = Dvbbs.Execute("select username from [dv_user] where   userid = "&faceid&"  ")
				If not rs.eof  Then
					usernames = rs(0)
				End If
				rs.close
				Set rs = Nothing
			End If		
		End If
		response.write "<tr><td class = forumRow height = 23><a href=""../uploadface/"&upname.name&""" target=_blank>"&upname.name&"</a></td>"
		response.write "<td align = right class = forumRowHighlight>"&usernames&"</td>"
		response.write "<td align = right class = forumRow>"& upname.size &"</td>"
		response.write "<td align = center class = forumRowHighlight>"& upname.datelastaccessed &"</td>"
		response.write "<td align = center class = forumRow>"& upname.datecreated &"</td>"
		response.write "<td align = center class = forumRowHighlight><a href = '?filename="&upname.name&"'>ɾ��</a></td></tr>"
	ElseIf i>page*pagesize Then
		Exit For
	End If
	usernames = ""
Next

End Sub 

'����ͷ��
Sub Delface()
	Server.ScriptTimeout = 999999
	On Error Resume Next
	Dim DllUserFace
	Dim Upfacepath, Newfilename
	If UpUserFaceFolder = "" Then
		Upfacepath = "../uploadFace/"
	Else
		Upfacepath = ".."&UpUserFaceFolder
	End If
	Dnum = 0
	DllUserFace = Request("filename")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	'ɾ������ͷ��������ļ�;
	If DllUserFace <> "" Then
		DllUserFace = Replace(DllUserFace,"..","")
		objFSO.DeleteFile(Server.MapPath(Upfacepath&DllUserFace))
		If Err Then
			Response.Write Err.Description
			Exit Sub
		End If
	End If
	Set uploadFolder = objFSO.GetFolder(Server.MapPath(Upfacepath))
	Set uploadFiles = uploadFolder.Files
	Filenum = uploadfiles.count
	i = 0
	For Each Upname In uploadFiles
		i = i + 1
		If i > 0 And i <= Filenum Then
			Upfilename = Lcase(Upfacepath&upname.name)
			'ȡ��ͷ����û���
			If Instr(upname.name,"_") Then
				Uid = Split(upname.name,"_")
				Faceid = Uid(0)
				If IsNumeric(Faceid) Then
					Set Rs = Dvbbs.Execute("SELECT Username, Userface FROM [Dv_User] WHERE Userid = " & Faceid)
					If Not (Rs.Eof And Rs.Bof) Then
						Usernames = Rs(0)
						Userface = Lcase(Trim(Rs(1)))
						If Instr(Userface,"|") > 0 Then Userface = Split(Userface,"|")(1)
						If Instr(Replace(Upfilename,"../",""),Userface) = 0 Then
							objFSO.DeleteFile(Server.MapPath(upfilename))
							Response.Write "ͷ���Ѹ���,�û�" & Usernames & "��ͷ���ļ���"& upfilename &"��ɾ��<br>"
							Response.Flush
							If Err Then
								Response.Write Err.Description
								Exit For
							End If
							Dnum = Dnum + 1
						End If
					Else
						objFSO.DeleteFile(Server.MapPath(upfilename))
						Response.Write "�û�ID��" & Faceid & "��ע��,�ļ���" & Upfilename &"��ɾ��<br>"
						Response.Flush
						If Err Then
							Response.Write Err.Description
							Exit For
						End If
						Dnum = Dnum + 1
					End If
					Set Rs = Nothing
				End If
			Else
			'����û���û�ID��ͷ���ļ�
				Sql = "SELECT Top 1 Userid From [Dv_User] WHERE Userface = '" & Upfilename & "'"
				Set Rs = Dvbbs.Execute(Sql)
				If Rs.Eof And Rs.Bof Then
					objFSO.DeleteFile(Server.MapPath(upfilename))
					Response.Write "�����ɾ���ļ���" & upfilename & "<br>"
					Response.Flush
					If Err Then
						Response.Write Err.Description
						Exit For
					End If
					Dnum = Dnum + 1
				Else
				'��Ϊ��ID��ͷ�� 2005-1-15 Dv.Yz
					Faceid = Rs(0)
					Newfilename = Upfacepath & Faceid & "_" & Upname.Name
					objFSO.Movefile ""&Server.MapPath(Upfilename)&"",""&Server.MapPath(Newfilename)&""
					If Not Err Then
						Dvbbs.Execute("UPDATE [Dv_User] Set UserFace = '"& Replace(Newfilename,"'","") & "' WHERE Userid = " & Faceid)
						Response.Write "��ͷ��" & Upfilename & " �Ѹ�Ϊ��" & Newfilename & "<br>"
						Response.Flush
						Dnum = Dnum + 1
					Else
						Response.Write Err.Description
						Exit For
					End If
				End If
				Set Rs = Nothing
			End If
		End If
	Next
	Response.Write " ������ "& dnum &" ���ļ�  "
End Sub

Sub foot()
Set uploadFolder = Nothing
Set uploadFiles = Nothing
%>
<tr><td colspan=6 class=forumRow height=30>
<%
If page>1 Then
	response.write "<a href=?page=1>��ҳ</a>&nbsp;&nbsp;<a href=""?page="& page-1 &""">��һҳ</a>&nbsp;&nbsp;"
Else
	response.write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
End If
If page<i/pagesize Then
	response.write "<a href=""?page="& page+1 &""">��һҳ</a>&nbsp;&nbsp;<a href=""?page="& pagenum &""">βҳ</a>"
Else
	response.write "��һҳ&nbsp;&nbsp;βҳ"
End If
%>
<input type="submit" value="����"></td><tr></form></table><br>
<% End Sub %>