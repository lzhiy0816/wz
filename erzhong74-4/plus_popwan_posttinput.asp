<!--#include file="Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dvbbs.LoadTemplates("")
Dvbbs.ErrType = 1	'ת������ʾ�����͵����Ĵ�����ʾҳ
Dvbbs.Head()

	If Not (Dvbbs.master Or Dvbbs.GroupSetting(70)="1") Then
		Dvbbs.AddErrcode(28)
		Dvbbs.ShowErr()
	End If

Dim action

Dim G_CurrentFolder,G_Msg
G_CurrentFolder = "Plus_popwan/DefaultInput/post/"

action = Request("action")

info()

Select Case action
	Case "add" : add()
	Case "edit" : edit()
	Case "save" : save()
	Case "del" : del()
	Case "demo" : demo()
	Case Else
		Call main()
End Select

Dvbbs.ActiveOnline()
Dvbbs.Footer()
Dvbbs.PageEnd()


Sub info()
%>
<br/>
<div class="tableborder" style="width:97%;">������<a href="?action=add">���ģ��</a> | <a href="plus_popwan_posttinput.asp">ģ���б�</a></div>
<br/>
<%
End Sub

Sub add()
%>
<form action="?act=add&action=save" method="post" name="theform">
<table class="tableborder" cellspacing="1">
	<tr><th colspan="2">�½�ģ��</th></tr>
	<tr>
		<td width="20%">�ļ���</td>
		<td width="80%"><input type="text" name="f_name" size="35"><font color="red">*</font>��ʽ��Ӣ��+����+�»���&nbsp;�磺abc_123</td>
	</tr>
	<tr>
		<td>�������</td>
		<td><input type="text" name="topic" size="35"></td>
	</tr>
	<tr>
		<td>��������</td>
		<td><textarea name="body" cols="80" rows="10"></textarea></td>
	</tr>
	<tr> 
		<td height="24">&nbsp;</td>
		<td><input type="submit" name="Submit" value="�ύ" class="button"></td>
	</tr>
</table>
</form>
<%
End Sub

Sub edit()
Dim f_name,Content
f_name = Replace(Request("f_name"),".html","")
If IsSafeParam(f_name,"^[a-zA-Z0-9_]+$") Then
	Content = Dvbbs.ReadTextFile(G_CurrentFolder&f_name&".html")
	Response.Write "<div style=""border:1px red solid;display : none;"">"&Content&"</div>" & vbNewLine
	%>
	<form action="?act=edit&action=save" method="post" name="theform">
	<table class="tableborder" cellspacing="1">
		<tr><th colspan="2">�༭ģ�壺<%=f_name%></th></tr>
		<tr>
			<td width="20%">�ļ���</td>
			<td width="80%"><input type="text" name="f_name" size="35" value="<%=f_name%>"><font color="red">*</font>��ʽ��Ӣ��+����+�»���&nbsp;�磺abc_123</td>
			<input type="hidden" name="f_oldname" value="<%=f_name%>"/>
		</tr>
		<tr>
			<td>�������</td>
			<td><input type="text" name="topic" size="35"></td>
		</tr>
		<tr>
			<td>��������</td>
			<td><textarea name="body" cols="80" rows="10"></textarea></td>
		</tr>
		<tr> 
			<td height="24">&nbsp;</td>
			<td><input type="submit" name="Submit" value="�ύ" class="button"></td>
		</tr>
	</table>
	</form>
	<script type="text/javascript" language="javascript">
	<!--
	document.theform.topic.value=document.getElementById('topic').innerHTML;
	document.theform.body.value=document.getElementById('body').innerHTML;
	//-->
	</script>
	<%
Else
	G_Msg = "���ݹ����Ĳ������淶���޷���ȡģ���ļ���"
	Response.Redirect "showerr.asp?ErrCodes=<li>"& G_Msg &"&action=NoHeadErr"
	Exit Sub
End If
End Sub

Sub save()
	Dim demo,demo_head,demo_content,demo_foot
	
	demo_head = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbNewLine
	demo_head = demo_head & "<html xml:lang=""zh-CN"" lang=""zh-CN"" xmlns=""http://www.w3.org/1999/xhtml"">" & vbNewLine
	demo_head = demo_head & "<head>" & vbNewLine
	demo_head = demo_head & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbNewLine
	demo_head = demo_head & "<title>demo</title>" & vbNewLine
	demo_head = demo_head & "<meta name=""MSSmartTagsPreventParsing"" content=""TRUE"" />" & vbNewLine
	demo_head = demo_head & "<meta http-equiv=""MSThemeCompatible"" content=""Yes"" />" & vbNewLine
	demo_head = demo_head & "</head>" & vbNewLine
	demo_head = demo_head & "<body>" & vbNewLine
	demo_head = demo_head & "<form id=""demoform"" name=""demoform"">" & vbNewLine

	demo_content = demo_content & "<div id=""topic"">"& Request("topic") &"</div>" & vbNewLine
	demo_content = demo_content & "<div id=""body"">"& Dvbbs.TextEnCode(Request("body")) &"</div>" & vbNewLine

	demo_foot = "</form>" & vbNewLine
	demo_foot = demo_foot & "</body>" & vbNewLine
	demo_foot = demo_foot & "</html>"

	demo = demo_head & demo_content & demo_foot
	savetofile demo
End Sub
Sub savetofile(content)
	Dim sLabelName, sLabelOldName, act
	act = Request("act")
	sLabelName=Replace(request("f_name"),".html","")
	sLabelOldName=Replace(request("f_oldname"),".html","")
	G_Msg=""
	If IsSafeParam(sLabelName,"^[a-zA-Z0-9_]+$") Then
		If act="add" Then
			If FileIsExist(G_CurrentFolder&sLabelName&".html") Then G_Msg="��ģ�����Ѵ��ڣ����޸�ģ�����������ύ��"
		Else
			If sLabelOldName<>"" And sLabelOldName<>sLabelName Then 
				If FileIsExist(G_CurrentFolder&sLabelName&".html") Then
					G_Msg="����ͼ�޸�ģ���������Ǹ�ģ�����Ѵ��ڣ����޸ĺ������ύ��"
				Else
					If Not FileReName(G_CurrentFolder&sLabelOldName&".html", sLabelName&".html") Then 
						G_Msg="����ͼ�޸�ģ����������û�гɹ���������Ȩ�޲�����"
					End If
				End If 
			End If 
		End If
		If ""=G_Msg Then 
			On Error Resume Next
			Dvbbs.WriteToFile G_CurrentFolder&sLabelName&".html", Content
			If Err Then
				Err.Description
				Response.End
				Err.Clear
				G_Msg="ģ�屣��ʧ�ܡ����������ļ��У�Plus_popwan/DefaultInput/post��������Ŀ¼û��д����޸�Ȩ�ޡ�"
			Else
				G_Msg="��ϲ��ģ�屣��ɹ���"
			End If
			Dvbbs.Dvbbs_suc(G_Msg)
		Else
			Response.Redirect "showerr.asp?ErrCodes=<li>"& G_Msg &"&action=NoHeadErr"
			Exit Sub
		End If
	Else
		G_Msg="ģ�������淶��ģ����ֻ������ĸ�����ֺ��»�����ɡ����޸ĺ������ύ��"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& G_Msg &"&action=NoHeadErr"
		Exit Sub
	End If
End Sub

Sub main()
	ListLabelFolder(G_CurrentFolder)
End Sub
Sub ListLabelFolder(sLabelPath)
Dim Fso, Folder, File, G_i
Set Fso	= CreateFSO()
sLabelPath = Server.MapPath(sLabelPath)

Set Folder = Fso.GetFolder(sLabelPath)
%>
<table class="tableborder" cellspacing="1">
	<tr><th colspan="2">ģ���б�</th></tr>
<%
	For Each File In Folder.Files
		G_i = G_i + 1
		response.write "<tr><td>" & File.name & "</td>" & vbNewLine
		Response.Write "<td><a href=""javascript:;"" onclick=""fillform('"&G_CurrentFolder&File.name&"','"&Replace(File.name,".html","")&"')"">��д</a> | <a href='?action=edit&f_name="&File.name&"'>�༭</a> | <a href='?action=del&f_name="&File.name&"' onclick='return confirm(""��ȷ��Ҫɾ��"&File.name&"ģ����ɾ��֮���ָܻ���"")'>ɾ��</a></td></tr>" & vbNewLine
	Next 
	Set File = Nothing 
	Set Fso	= Nothing 
	If 0=G_i Then
		response.write "<tr><td colspan=""2"">��δ��ӣ�</td></tr>"
	End If 
%>
</table>

<SCRIPT LANGUAGE="JavaScript">
<!--
	function fillform(path,file){
		var a=document.createElement("iframe");
		a.id = file;
		a.name = file;
		a.scrolling = "no";
		a.src = path;
		a.style.display = 'none';
		document.body.appendChild(a);
		alert('������');
		var b=document.getElementById(file);
		//b.contentWindow.document.location.reload();
		var topic = b.contentWindow.document.getElementById('topic');
		var body = b.contentWindow.document.getElementById('body');
		//��д
		parent.document.Dvform.topic.value=topic.innerHTML;
		parent.dvtextarea.clear();
		parent.dvtextarea.insert(body.innerHTML);
		wndClose();
	};
	function wndClose(){
		try{
			parent.DvWnd.close();
		}
		catch(e){
			window.close();
		}
	};
//-->
</SCRIPT>
<%
End Sub

'ɾ���ļ�
Sub del()
	Dim Fso, sLabelName, sRealPath
	sLabelName=Replace(request("f_name"),".html","")
	If IsSafeParam(sLabelName,"^[a-zA-Z0-9_]+$") Then 
		On Error Resume Next 
		sRealPath=Server.MapPath(G_CurrentFolder&sLabelName&".html")
		Set Fso=CreateFSO()
		If Fso.FileExists(sRealPath) Then
			Fso.DeleteFile sRealPath,True
			If Err Then
				Err.Clear
				G_Msg="��ɾ���ļ�ʱ�������󣬿�����û���㹻��Ȩ�ޡ��뵽�ռ����ֶ�ɾ�����ļ���"
			Else
				G_Msg="�ɹ�ɾ���ļ���"&sLabelName&""
			End If
		Else
			G_Msg="�ļ�û���ҵ��������Ѿ���ɾ��������û���㹻��Ȩ�ޡ�"
		End If 
		Set Fso=Nothing
	Else
		G_Msg = "���ݹ�����·����Ϊ��ȫԭ�򱻽�ֹ���뵽�ռ����ֶ�ɾ�����ļ���"
		Response.Redirect "showerr.asp?ErrCodes=<li>"& G_Msg &"&action=NoHeadErr"
		Exit Sub
	End If
	Dvbbs.Dvbbs_suc(G_Msg)
End Sub


Function CreateFSO()
	On Error Resume Next 
	Set CreateFSO = Dvbbs.iCreateObject("Scripting.FileSystemObject")
	If Err Then 
		Err.Clear
		response.write "���Ŀռ䲻֧��FSO������FSO���������ڰ�ȫԭ�򱻸��Ĺ�������ռ�����ϵ��"
		response.End 
	End If
End Function

Function FileIsExist(Path)
	Dim Fso:FileIsExist=False 
	On Error Resume Next 
	Set Fso=CreateFSO()
	If Fso.FileExists(Server.MapPath(Path)) Then FileIsExist=True 
	Set Fso=Nothing 
End Function

Function FileReName(Path,NewName)
	Dim Fso,File
	FileReName=False 
	On Error Resume Next 
	Set Fso=CreateFSO()
	Set File=Fso.GetFile(Server.MapPath(Path))
	File.name=NewName
	Set File=Nothing 
	Set Fso=Nothing 
	FileReName=True 
End Function

Function IsSafeParam(Path,Param)
	Dim re
	IsSafeParam=False 
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern=Param
	IsSafeParam=re.Test(Path)
	Set Re=Nothing
End Function
%>