<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Dim admin_flag,CssList,StyleConn
Head()
admin_flag=",23,"
CheckAdmin(admin_flag)
Select Case Request("action")
	Case "registerTemplate"
		RegisterTemplate()
	Case "logOutTemplate"
		LogOutTemplate()
	Case Else
		Main()
End Select
Footer()

Sub Main()
%>
<table border="0" cellspacing="1" cellpadding="5" align="center" width="100%">
	<tr>
		<th colspan="3" style="text-align:center;" id="TableTitleLink">ģ��ע���ע������</th>
	</tr>
	<tr>
		<td class="forumHeaderBackgroundAlternate">
		ע��ģ��˵����<br/>
		�� ����Ҫע���ģ���Ŀ¼�ϴ�����̳��Ŀ¼�µ�<font color="#FF0000">Resource</font>Ŀ¼�У�ģ��Ŀ¼���Ʋ���ʹ�ú��֣���<br/>
		�� ��д���桰ע��ģ�塱�������Ϣ�ύ���ɡ�<br/>
		�� ע�⣺<font color="#FF0000">ģ��Ŀ¼ֻдģ����ļ������ƣ��磺Template_1��������ҪResource/ǰ׺��</font><br/>
		<div id="formRegTemp">
		<form name="formRegTemplate" method="post" action="?action=registerTemplate" onsubmit="return checkform(this);" target="hiddenframe">
			ģ�����ƣ�<input type="text" name="TemplateName" size="15"/>
			ģ��Ŀ¼��<input type="text" name="TemplateFolder" size="15"/>
			<input type="submit" name="register" value="ע��">
			<br/>ע����ģ���������Ϊ�գ�����ģ��Ŀ¼��Ϊģ�����ƣ�
			<br/>&nbsp;&nbsp;&nbsp;&nbsp;��ģ��Ŀ¼����Ϊ�ա�
			<div id="ErrorInfo" style="display:none;"></div>
		</form>
		</div>
		</td>
	</tr>
	<tr>
		<td class="forumHeaderBackgroundAlternate">
		ע��ģ��˵����<br/>
		�ٵ����Ӧģ��ġ�ע������ť����ע��ģ�壻<br/>
		�ڽ���ע����ģ���Ŀ¼����̳��Ŀ¼�µ�<font color="#FF0000">Resource</font>Ŀ¼���Ƴ���<br/>
		<form name="formLogOutTemplate" method="post" action="?action=logOutTemplate" onsubmit="return logOutTemplate(0,'','');">
		<input type="hidden" name="id" value=""/>
		<input type="hidden" name="folder" value=""/>
		<table width="40%">
			<tr>
				<td>ģ������</td>
				<td>ģ��Ŀ¼</td>
				<td>����</td>
			</tr>
		<%
		Dim Rs,SQL
		Set Rs=Dvbbs.Execute("Select * From Dv_Templates")
		Do While Not Rs.Eof
			Response.write "<tr><td>"&Rs(1)&"</td><td>"&Rs(2)&"</td><td><input type=""button"" name=""LogOut"" value=""ע��"" onclick=""logOutTemplate("&Rs(0)&",'"&Rs(1)&"','"&Rs(2)&"')""/></td></tr>"&chr(10)
			Rs.MoveNext
		Loop
		%>
		</table>
		</form>
		</td>
	</tr>
</table>
<iframe style="border:0px;width:0px;height:0px;" src="" name="hiddenframe" id="hiddenframe"></iframe>
<script languange="javascript">
	function checkform(theForm){
		if (''==theForm.TemplateFolder.value){
			document.getElementById('ErrorInfo').innerHTML='<img src="../skins/default/images/note_error.gif"/><font color="red">ģ��Ŀ¼����Ϊ�ա�</font>';
			document.getElementById('ErrorInfo').style.display='block';
			theForm.TemplateFolder.focus();
			return false;
		}
		if ('none'!=document.getElementById('ErrorInfo').style.display)
			document.getElementById('ErrorInfo').style.display='none';
		return true;
	}
	function logOutTemplate(id,templateName,templateFolder){
		if (0 != id){
			if (confirm('ȷ��Ҫע��ģ�塰'+templateName+'����')){
				document.formLogOutTemplate.id.value=id;
				document.formLogOutTemplate.folder.value=templateFolder;
				document.formLogOutTemplate.submit();
			}		
		}
		return false;
	}
</script>
<%
End Sub

Sub RegisterTemplate()
	Dim TemplateName,TemplateFolder
	Dim Rs
	TemplateName = Dvbbs.CheckStr(Request.Form("TemplateName"))
	TemplateFolder = Dvbbs.CheckStr(Request.Form("TemplateFolder"))
	If TemplateFolder="" Then
		Response.write "<script language=""javascript"">"
		Response.write "parent.document.getElementById('ErrorInfo').innerHTML='<img src=""../skins/default/images/note_error.gif""/><font color=""red"">ģ��Ŀ¼����Ϊ�ա�</font>';"
		Response.write "parent.document.getElementById('ErrorInfo').style.display='block';"
		Response.write "</script>"
		Response.end
	End If
	If TemplateName="" Then TemplateName=TemplateFolder
	Set Rs=Dvbbs.Execute("Select * From Dv_Templates Where Folder='"&TemplateFolder&"'")
	If Not Rs.Eof Then
		Response.write "<script language=""javascript"">"
		Response.write "parent.document.getElementById('ErrorInfo').innerHTML='<img src=""../skins/default/images/note_error.gif""/><font color=""red"">ģ��Ŀ¼��"&TemplateFolder&"���Ѿ�ע����������ظ�ע�ᡣ</font>';"
		Response.write "parent.document.getElementById('ErrorInfo').style.display='block';"
		Response.write "</script>"
		Response.end
	End If
	Rs.Close:Set Rs=Nothing
	Dvbbs.Execute("Insert Into Dv_Templates(Type,Folder) Values('"&TemplateName&"','"&TemplateFolder&"')")
	Response.write "<script language=""javascript"">"
	Response.write "parent.document.getElementById('formRegTemp').innerHTML='<img src=""../skins/default/images/note_ok.gif""/>"&TemplateName&" ģ��ע��ɹ���<a href=""Template_RegAndLogout.asp"">����</a>����ע��ģ��';"
	Response.write "</script>"
	Dvbbs.Loadstyle()
End Sub

Sub LogOutTemplate()
	Dim ID,TemplateFolder,rs
	ID = Dvbbs.CheckNumeric(Request.Form("id"))
	TemplateFolder = Request.Form("folder")
	Dim count1
	Set Rs=Dvbbs.Execute("Select count(*) From Dv_Templates")
	count1=Rs(0)
	Rs.close
	Set Rs=Nothing
	If(count1>1)Then
		Dvbbs.Execute("Delete From Dv_Templates Where ID="&ID)
		Dvbbs.Loadstyle()
		Response.Write "<script language='javascript'>"
		Response.Write "alert('ģ��ע���ɹ������¼ftp�Ѹ�ģ���Ӧ���ļ���"&TemplateFolder&"ɾ����');"
		Response.Write "self.location='Template_RegAndLogout.asp';"
		Response.Write "</script>"
	Else
		Response.Write "<script language='javascript'>"
		Response.Write "alert('Ψһ��һ��ģ���ǲ����Ա�ע����');"
		Response.Write "self.location='Template_RegAndLogout.asp';"
		Response.Write "</script>"
	End if
	Dvbbs.Loadstyle()
End Sub
%>