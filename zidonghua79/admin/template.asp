<!--#include file="../conn.asp"-->
<!--#include file="inc/const.asp"-->
<%
Head()
Dim admin_flag
admin_flag=",20,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<br /><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br /><li>��û�й���ҳ���Ȩ�ޡ�</li>"
	dvbbs_error()
End If
Dim action,SkinID,StyleID
StyleID=request("StyleID")
action=Request("action")
If Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']") Is Nothing Then
	If Not Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@id") Is Nothing Then
		StyleID=Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style/@id").text 
	Else
		Response.Write "ģ�������޷���ȡ,��������µ���"
		Response.End
	End If
End If
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
Response.Write "<tr>"
Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=""2"" height=""25"">��̳ģ�����"
Response.Write "</th>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td class=""forumRowHighlight"" colspan=""2"">"
Response.Write "<p><b>ע��</b>��<br />��������������½����޸�ģ�壬���Ա༭��̳���԰��ͷ�񣬿����½�ģ��ҳ�棬����ʱ�밴�����ҳ����ʾ������д����Ϣ��<br />����̳��ǰ����ʹ�õ�Ĭ��ģ�岻��ɾ��<br />������޸ķ�ģ��ҳ�����ƻ�ɾ����ģ��ҳ�����ڹر���̳֮�����,������ܻ�Ӱ����̳����."
Response.Write "</td>"
Response.Write "</tr>"
Response.Write "<tr>"
Response.Write "<td class=""forumRowHighlight"" width=""20%"" height=""25"" align=""left"">"
Response.Write "<b>��̳ģ�����ѡ��</b></td>"
Response.Write "<td class=""forumRowHighlight"" width=""80%""><a href=""template.asp"">ģ�������ҳ</a>"
Response.Write " | <a href=""http://bbs.dvbbs.net/loadtemplates.asp"" target = ""_blank"">��ȡ�ٷ�ģ������</a></td>"
Response.Write "</tr>"
Response.Write "</table>"
Select Case action
	Case "edit"
		Call Edit() 
	Case "manage"
		If Request("mostyle")="�� ��" Then
			Main()
		ElseIf Request("mostyle") = "ɾ ��" Then
			DelStyle()
		End If
	Case "saveedit"
		Call Saveedit()
	Case "addpage"
		addpage()
	Case "addstyle"
		addstyle()
	Case "Edit_Main"
		Edit_Main()
	Case "Save_Main"
		Save_Main()
	Case "rename"
		rename()
	Case "editcss"
		editcss()
	Case "savecss"
		savecss()
	Case "editmain"
		editmain()
	Case "savemain"
		Savemain() 
	Case "ghost"
		ghost()
	Case "delpage"
		delpage()
	Case "pagerename"
		pagerename()
	Case Else
		Main()
End Select

footer()
Sub Main()
	Response.Write "<p></p>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=""2"" height=""25"">��ǰ��̳ģ�����"
	Response.Write "</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<form method=post action=""?action=manage"">"
	Response.Write "<td class=""forumRowHighlight"" height=40 align=left>"
	Response.Write "��ѡ�����ģ�壺 "
	'����ϵͳ��������ȡ������ģ�����ƺ�ID
	Dim Templateslist
	Response.Write "<select name=""StyleID"" size=""1"">"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@stylename").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "<input type=submit value=""�� ��"" name=""mostyle"">&nbsp;&nbsp;&nbsp;"
	Response.Write "<input type=submit value=""ɾ ��"" name=""mostyle"">"
	Response.Write "<br /><br /><b>˵����</b>ɾ��������ɾ����ģ���������ݣ����á�"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=addpage"">"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
	Response.Write "�½���ģ��ҳ�棺&nbsp;"
	Response.Write "<input type=text size=25 name=""StylePageName"">&nbsp;&nbsp;"
	Response.Write "<input type=submit name=submit value=""�� ��""> "
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=right colspan=2>"
	Response.Write "��������ASPҳ��������Ҫ������׺�����½�����ҳ��ģ����Ǹ�ҳ���ģ����Դ������ͼƬ�����ԡ���񣩡�"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<th height=25 colspan=2>"
	Response.Write Application(Dvbbs.CacheName &"_style").documentElement.selectSingleNode("style[@id='"& StyleID &"']/@stylename").text &"����ģ����Դ����</th></tr><tr><td height=25 class=""bodytitle"" colspan=2>"
	Response.Write "ͨ����˵����ҳ��ģ�������̳��ÿ��ҳ��ķ��ģ�壬���������ֶ������ֶε���������Ϊ��Page_ҳ��������Ҫ��׺��"
	Response.Write "</td>"
	Response.Write "</tr>"
	Set Rs=Dvbbs.Execute("Select top 1 * From Dv_Style ")
	For i= 1 to Rs.Fields.Count-1
		If i = 12 Then
			Response.Write "<tr onmouseover=""this.style.backgroundColor='#B3CFEE';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"">"
			Response.Write "<td height=""25"">"
			Response.Write "<li>"
			Response.Write "��ҳ��ģ��<a href=#>(page_admin)</a>&nbsp;&nbsp;</td><td height=""25"" align=""left"">"
			Response.Write "�༭��ģ�飺"
			Response.Write "<a href=""?action=Edit_Main&stype=1&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>���԰�</a> <a href=""?action=Edit_Main&stype=1&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>��</font></a> | <a href=""?action=Edit_Main&stype=2&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>ͼƬ</a> <a href=""?action=Edit_Main&stype=2&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new"""
			Response.Write "><font color=blue>��</font></a> | <a href=""?action=Edit_Main&stype=3&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>������</a> <a href=""?action=Edit_Main&stype=3&page="
			Response.Write "page_admin"
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>��</font></a>"
			Response.Write "</td>"
			Response.Write "</tr>"
		End If
		If i> 20 Then
			Response.Write "<form method=post action=""?action=pagerename"">"
		End If
		Response.Write "<tr onmouseover=""this.style.backgroundColor='#B3CFEE';this.style.color='red'"" onmouseout=""this.style.backgroundColor='';this.style.color=''"">"
		Response.Write "<td height=25 align=left>"
		Response.Write "<li>"
		Select Case i
			Case 1
				Response.Write "��ǰģ��CSS����<a href=#>"
			Case 2
				Response.Write "��ǰģ����ģ��<a href=#>"
			Case Else 
				Response.Write "��ҳ��ģ��<a href=#>"
		End Select
		
		If i> 20 Then
			
			Response.Write "</a>&nbsp;&nbsp;page_"
			Response.Write "<input type=text size=15 name=newpagename value="&Replace(Rs(i).Name,"page_","")&">&nbsp;&nbsp;"
			Response.Write "<input type=hidden size=15 name=oldpagename value="&Replace(Rs(i).Name,"page_","")&">"		
			Response.Write "<input type=submit value=""��ģ��ҳ�����"" name=""mo"" title=""�޸����ƺ��ύ"">"
		Else
			Response.Write "("&Rs(i).Name&")</a>&nbsp;&nbsp;"		
		End If		
		Response.Write "</td>"
		Response.Write "<td height=""25"" align=""left"">"
		If i=3 Then
			Response.Write "�༭��ģ�飺"
			Response.Write "<a href=""bbsface.asp?Stype=1&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>��������</a> <a href=""?action=edit&stype=1&page="					
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""></a> | <a href=""bbsface.asp?Stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>��������</a> <a href=""?action=edit&stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new"""
			Response.Write "></a> | <a href=""bbsface.asp?Stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>�û�ͷ��</a> <a href=""?action=edit&stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""></a>"		
		ElseIf i>1 Then 
			Response.Write "�༭��ģ�飺"
			Response.Write "<a href=""?action=edit&stype=1&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>���԰�</a> <a href=""?action=edit&stype=1&page="					
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>��</font></a> | <a href=""?action=edit&stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>ͼƬ</a> <a href=""?action=edit&stype=2&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new"""
			Response.Write "><font color=blue>��</font></a> | <a href=""?action=edit&stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write """>������</a> <a href=""?action=edit&stype=3&page="
			Response.Write Rs(i).Name
			Response.Write "&StyleID="
			Response.Write StyleID
			Response.Write "#new""><font color=blue>��</font></a>"
			If i=2 Then 
				Response.Write " | <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">��������</a>"		
			End If
		ElseIf i=1 Then
			Response.Write "�༭��ģ�飺"
			Response.Write "<a href=""?action=editcss&StyleID="&StyleID&""">�޸�CSS��ʽ</a>"	 
		End If
		If i >20 Then 
			Response.Write "&nbsp;&nbsp;<a href=""?action=delpage&StylePageName="&Rs(i).Name&""" title=""ע�⣬ɾ���󲻿ɻָ�""> ɾ����ģ��ҳ�� </a>"		
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		If i>20 Then
			Response.Write "</form>" 
		End If
	Next
	Response.Write "</table><p></p>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>��̳ģ�����"
	Response.Write "</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=addstyle"">"
	Response.Write "<td class=""forumRowHighlight"" height=40 align=left width=45% >"
	Response.Write "�½�ģ�壺&nbsp;"
	Response.Write "<input type=text size=25 name=""StyleName"">&nbsp;&nbsp;"
	Response.Write "<input type=submit value=""�� ��"" name=""mostyle"">&nbsp;&nbsp;��дģ����"
	Response.Write "<br /><br /><b>˵����</b>�½�ģ�彫�ѵ�ǰ��̳Ĭ��ģ�����ݸ��Ƶ��µ�ģ����"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=ghost"">"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>Դģ�壺"
	Response.Write "<select name=""OlDStyleID"" size=1>"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@stylename").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;Ŀ��ģ�壺"
	Response.Write "<select name=""newStyleID"" size=1>"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@stylename").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;<input type=submit name=submit value=""ģ���¡"">"
	Response.Write "<br /><br /><b>˵����</b>ģ���¡����Դģ�����ݸ���Ŀ��ģ���е���������"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<FORM METHOD=POST ACTION=""?action=rename"">"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
	Response.Write "<select name=""StyleID"" size=1>"
	For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<option value="""& Templateslist.selectSingleNode("@id").text &""""
		If CLng(Templateslist.selectSingleNode("@id").text) = CLng(StyleID) Then 
			Response.Write " selected"
		End If 
		Response.Write ">"&Templateslist.selectSingleNode("@stylename").text &"</option>"
	Next 
	Response.Write "</select>"
	Response.Write "&nbsp;&nbsp;"
	Response.Write "����Ϊ��<input type=text size=20 name=""StyleName"" value="""
	Response.Write """>&nbsp;&nbsp;"
	Response.Write "<input type=submit name=submit value=""�޸�"">"
	Response.Write "</td>"
	Response.Write "</FORM>"
	Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
	Response.Write "&nbsp;"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table><br />"
	Rs.Close
	Set Rs=Nothing
End Sub
Sub Edit()
	Dim Page,mystr
	Dim TempStr,TemplateStr,stype
	Dim TempStyleHelp,StyleHelpValue
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<br /><li>�������ʽ����"
		Dvbbs_error()
	End If
	If Not IsTruePage(page) Then
		Errmsg=ErrMsg + "<br /><li>Ҫ�༭��ҳ��ģ���ֶ���δ������"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_StyleHelp] where ID=1")
	TempStr=Split(Rs(2),"@@@")
	Select Case stype
		Case 1
			TempStyleHelp=Split(TempStr(1),"|||")
		Case 2
			TempStyleHelp=Split(TempStr(2),"|||")
		Case 3
			TempStyleHelp=Split(TempStr(0),"|||")
	End Select
	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_Style] Where ID="&StyleID)
	TempStr=Split(Rs(2),"@@@")
	Select Case stype
		Case 1
			TemplateStr=Split(TempStr(1),"|||")
		Case 2
			TemplateStr=Split(TempStr(2),"|||")
		Case 3
			TemplateStr=Split(TempStr(0),"|||")
	End Select
	Response.Write "<form name=""template"" action=""?action=saveedit&page="&page&"&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=3 height=25>"
	Response.Write Rs(1)
	Response.Write "��ҳ��ģ��("
	Response.Write page
	Response.Write ")"
	Response.Write "<input Type=""hidden"" name=""dvbbs"" value=""OK!"">"
	Select Case stype
		Case 1
			Response.Write "���԰�"
			mystr="template.Strings"
			If page="main_style" Then mystr="Dvbbs.lanstr"
		Case 2
			Response.Write "ͼƬ��Դ(��ǰĬ��·��{$PicUrl}Ϊ��"&Dvbbs.Forum_PicUrl&")"
			mystr="template.pic"
			If page="main_style" Then mystr="Dvbbs.mainpic"
		Case 3
			Response.Write "������"
			mystr="template.html"
			If page="main_style" Then mystr="Dvbbs.mainhtml"
	End Select
	
	Response.Write "����</th></tr>"
	If TemplateStr(Ubound(TemplateStr))="" Then TemplateStr(Ubound(TemplateStr))="del"
	For i=0 To Ubound(TemplateStr)
		If i<ubound(TempStyleHelp) Then
			StyleHelpValue=TempStyleHelp(i)
		Else
			StyleHelpValue="//"
		End IF
		Response.Write "<tr><td class=""forumRowHighlight"" width=20% height=40 align=left>"
		Response.Write mystr&"("&i&")"
		Response.Write "<br /><a href=""#"" onclick=""rundvscript(t"&i&",'page="&page&"&index="&i&"&stype="&stype&"');"" title=""�������ȡ�ⲿ��ģ��Ĺٷ�����"">��ȡ�ٷ�����</a>"
		Response.Write "</td>"		
		Response.Write "<td class=""forumRowHighlight"" width=80% height=25 align=left>"
		Select Case stype
			Case 1
				If LenB(TemplateStr(i))>70 Then
				Response.Write "<textarea name=""TemplateStr"" id=""t"&i&"""  cols=""100"" rows=""3"">"
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write "</textarea>"
				Else
				Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t"&i&""" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write """ size=50>"
				End If
				Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r"&i&""" value="""&StyleHelpValue&""">"
				Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""images/help.gif"" border=0 title=""������Ĺ��������""></a>"
			Case 2
				Response.Write "<input Type=""text"" name=""TemplateStr"" id=""t"&i&""" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write """ size=20> "
				If server.htmlencode(TemplateStr(i))<>"" And (Instr(server.htmlencode(TemplateStr(i)),".gif") or Instr(server.htmlencode(TemplateStr(i)),".jpg")) Then Response.Write "<img src="&server.htmlencode(Replace(TemplateStr(i),"{$PicUrl}",MyDbPath&Dvbbs.Forum_PicUrl))&"  border=0>"	
			Case 3
				If page="main_style"  And i=0 Then 
					Response.Write "<input type=hidden name=""TemplateStr"" value="""
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write """>"
					Response.Write "���ֶ����ڻ������ã�  <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">�������޸Ļ�������</a>"
					Response.Write "</td><td class=""forumRowHighlight"">"
					Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""images/help.gif"" border=0 title=""������Ĺ��������""></a>"
				Else
					
					Response.Write "<textarea name=""TemplateStr"" id=""t"&i&""" cols=""100"" rows=""5"">"
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write "</textarea>"
					Response.Write "</td><td class=""forumRowHighlight""><a href=""javascript:admin_Size(-5,'t"&i&"')""><img src=""images/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(5,'t"&i&"')""><img src=""images/plus.gif"" unselectable=""on"" border='0'></a>"
					Response.Write "<img src=images/viewpic.gif onclick=runscript(t"&i&")>"
					Response.Write "<a href=# onclick=""helpscript(r"&i&");return false;"" class=""helplink""><img src=""images/help.gif"" border=0 title=""������Ĺ��������""></a> "		
				End If
				Response.Write "<INPUT TYPE=""hidden"" NAME=""ReadME"" id=""r"&i&""" value="""&StyleHelpValue&""">"
			End Select
			
		Response.Write "</td></tr>"
	Next
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=""3"">"
	Response.Write "</td></tr>"
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""�� ��"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" colspan=2 align=""center"">"
	Response.Write "<input type=""submit"" name=""B1"" value=""�� ��"">"
	Response.Write "</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan=3 Class=""forumRowHighlight"">"
	Response.Write "<br /><li>��Ҫ��ʾ,ģ���к�XSLT�����,�޸ı����ϸ���XML�﷨��׼."
	Response.Write "<br /><li>ģ��༭���������������ֶΣ����ڶ�Ӧ���ı���������""del""����ôģ�����ݵ���žͻ�ǰ�ơ�"
	Response.Write "<br /><li>�������ı�ģ�����ݵ����,���Ѹ���Ŀ���������,��ֻ��Ҫ��������ա�"
	Response.Write "</td></tr>"
	Response.Write "</table><p></p>"
	Response.Write "</form>"
	Rs.Close
	Set Rs=Nothing
End Sub
Sub SaveEdit()
	If Request("dvbbs")<>"OK!" Then
		Errmsg=ErrMsg + "<br /><li>���ύ�˷Ƿ�����"
		Dvbbs_error()
		Exit Sub
	End If
	Dim Page
	Dim TempStr,TemplateStr,stype
	Dim TempStyleHelp,StyleHelpValue
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<br /><li>�������ʽ����"
		Dvbbs_error()
	End If
	If Not IsTruePage(page) Then
		Errmsg=ErrMsg + "<br /><li>Ҫ�༭��ҳ��ģ���ֶ���δ������"
		Dvbbs_error()
	End If
	'ģ����,���»���.	
	If stype="3" Then
	Select Case Request("page")
		Case "page_dispbbs"
			TemplateStr=Request.form("TemplateStr")(1)
			Set TempStr=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			If Not TempStr.Loadxml(TemplateStr) Then
				Errmsg=ErrMsg + "��̳��ҳģ��template.html(0)δ��ͨ��XMLУ��,�����±༭�޸�,ȷ������."
				Set TempStr=Nothing
				Dvbbs_error()
				Exit Sub
			End If
		Case "page_index"
			TemplateStr=Request.form("TemplateStr")(1)
			Set TempStr=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			If Not TempStr.Loadxml(TemplateStr) Then
				Errmsg=ErrMsg + "��̳��ҳģ��template.html(0)δ��ͨ��XMLУ��,�����±༭�޸�,ȷ������."
				Set TempStr=Nothing
				Dvbbs_error()
				Exit Sub
			End If
			TemplateStr=Request.form("TemplateStr")(2)
			Set TempStr=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If Not TempStr.Loadxml(TemplateStr)  Then
				Errmsg=ErrMsg + "��̳��ҳģ��template.html(1)δ��ͨ��XMLУ��,�����±༭�޸�,ȷ������."
				Set TempStr=Nothing
				Dvbbs_error()
				Exit Sub
			End If
			TemplateStr=Request.form("TemplateStr")(4)
			Set TempStr=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If Not TempStr.Loadxml(TemplateStr)  Then
				Errmsg=ErrMsg + "��̳��ҳģ��template.html(3)δ��ͨ��XMLУ��,�����±༭�޸�,ȷ������."
				Set TempStr=Nothing
				Dvbbs_error()
				Exit Sub
			End If
		Case "page_query"
			TemplateStr=Request.form("TemplateStr")(1)
			Set TempStr=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			If Not TempStr.Loadxml(TemplateStr) Then
				Errmsg=ErrMsg + "��̳��ҳģ��template.html(0)δ��ͨ��XMLУ��,�����±༭�޸�,ȷ������."
				Set TempStr=Nothing
				Dvbbs_error()
				Exit Sub
			End If
		Case "main_style"
			TemplateStr=Request.form("TemplateStr")(23)
			Set TempStr=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
			If Not TempStr.Loadxml(TemplateStr) Then
				Errmsg=ErrMsg + "��̳��ҳģ��Dvbbs.mainhtml(22)δ��ͨ��XMLУ��,�����±༭�޸�,ȷ������."
				Set TempStr=Nothing
				Dvbbs_error()
				Exit Sub
			End If
		End Select
	End If
	'��ȡ���е�����
	TemplateStr=""
	For Each TempStr in Request.form("TemplateStr")
		If LCase(TempStr)<>"del" Then 
			TemplateStr=TemplateStr&Replace(TempStr,"|||","")&"|||"
		End If
	Next
	TemplateStr=Dvbbs.checkStr(Replace(TemplateStr,"@@@",""))
	If Trim(TemplateStr)="" Then 
		TemplateStr="|||"
	End If

	'��ȡ���е�����
	StyleHelpValue=""
	For Each TempStyleHelp in Request.form("ReadME")
		If TempStyleHelp<>"" Then 
			StyleHelpValue=StyleHelpValue&TempStyleHelp&"|||"
		End If
	Next
	If Trim(StyleHelpValue)="" Then 
		StyleHelpValue="|||"
	Else
		StyleHelpValue=Dvbbs.checkStr(StyleHelpValue)
	End If

	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_Style] Where ID="&StyleID)
	TempStr=Split(Dvbbs.checkStr(Rs(2)),"@@@")
	Select Case stype
		Case 1
			TemplateStr=TempStr(0)&"@@@"&TemplateStr&"@@@"&TempStr(2)
		Case 2
			TemplateStr=TempStr(0)&"@@@"&TempStr(1)&"@@@"&TemplateStr
		Case 3
			TemplateStr=TemplateStr&"@@@"&TempStr(1)&"@@@"&TempStr(2)
	End Select

	Set Rs=Dvbbs.Execute("Select ID,StyleName,"&page&" From [Dv_StyleHelp] Where ID=1")
	TempStr=Split(Dvbbs.checkStr(Rs(2)),"@@@")
	Select Case stype
		Case 1
			StyleHelpValue=TempStr(0)&"@@@"&StyleHelpValue&"@@@"&TempStr(2)
		Case 2
			StyleHelpValue=TempStr(0)&"@@@"&TempStr(1)&"@@@"&StyleHelpValue
		Case 3
			StyleHelpValue=StyleHelpValue&"@@@"&TempStr(1)&"@@@"&TempStr(2)
	End Select
	Rs.close:Set Rs=Nothing
	Dvbbs.Execute("update [Dv_Style] set "&page&"='"&TemplateStr&"' Where ID="&StyleID)
	Dvbbs.Execute("update [Dv_StyleHelp] set "&page&"='"&StyleHelpValue&"' Where ID=1")
	If stype="3" Then
	Select Case Request("page")
		Case "page_dispbbs"
				Application.Lock
				Application.Contents.Remove(Dvbbs.CacheName & "_dispbbsemplate_"& Request("StyleID"))
				Application.unLock
		Case "page_index"
			Application.Lock
			Application.Contents.Remove(Dvbbs.CacheName & "_listtemplate_"& Request("StyleID"))
			Application.Contents.Remove(Dvbbs.CacheName & "_indextemplate_"& Request("StyleID"))
			Application.Contents.Remove(Dvbbs.CacheName & "_shownews_"&Request("StyleID"))
			Application.unLock
		Case "page_query"
				Application.Lock
				Application.Contents.Remove(Dvbbs.CacheName & "_querytemplate_"& Request("StyleID"))
				Application.unLock
		Case "main_style"
			RestoreBoardCache()
		Case Else
		End Select
	End If
	Select Case stype
		Case 1
			Dv_suc(page&"���԰��޸ĳɹ�!")
		Case 2
			Dv_suc(page&"ͼƬ��Դ�޸ĳɹ�!")
		Case 3
			Dv_suc(page&"�������޸ĳɹ�!")		
	End Select
	'���»��档�˴�����ģ�����ݱ仯��ʱ����Ҫ���µĴ��롣����©��������������ӡ�
	Dvbbs.Loadstyle()
End Sub
'��̨ģ��༭
Sub Edit_Main()
	Dim Page,mystr
	Dim TempStr,TemplateStr,stype
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<br /><li>�������ʽ����"
		Dvbbs_error()
	End If
	If page<>"page_admin" Then
		Errmsg=ErrMsg + "<br /><li>Ҫ�༭��ҳ��ģ���ֶ���δ������"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("Select H_ID,H_Title,H_Content From [Dv_help] Where H_ID=1")
	TempStr=Split(Rs(2),"@@@")
	Select Case stype
		Case 1
			TemplateStr=Split(TempStr(1),"|||")
		Case 2
			TemplateStr=Split(TempStr(2),"|||")
		Case 3
			TemplateStr=Split(TempStr(0),"|||")
	End Select
	Response.Write "<form action=""?action=Save_Main&page="&page&"&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=""3"" height=""25"">"
	Response.Write Rs(1)
	Response.Write "��ҳ��ģ��("
	Response.Write page
	Response.Write ")"
	Select Case stype
		Case 1
			Response.Write "���԰�"
			mystr="template.Strings"
			If page="main_style" Then mystr="Dvbbs.lanstr"
		Case 2
			Response.Write "ͼƬ��Դ"
			mystr="template.pic"
			If page="main_style" Then mystr="Dvbbs.mainpic"
		Case 3
			Response.Write "������"
			mystr="template.html"
			If page="main_style" Then mystr="Dvbbs.mainhtml"
	End Select
	
	Response.Write "����</th></tr>"
	For i=0 To Ubound(TemplateStr)
		Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=left>"
		Response.Write mystr&"("&i&")"
		Response.Write "</td>"		
		Response.Write "<td class=""forumRowHighlight"" height=25 align=left>"
		Select Case stype
			Case 1
				If LenB(TemplateStr(i))>70 Then
				Response.Write "<textarea name=""TemplateStr"" cols=""100"" rows=""3"">"
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write "</textarea>"
				Else
				Response.Write "<input Type=""text"" name=""TemplateStr"" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write """ size=50>"
				End If
			Case 2
				Response.Write "<input Type=""text"" name=""TemplateStr"" value="""
				Response.Write server.htmlencode(TemplateStr(i))
				Response.Write ""
				Response.Write """ size=20>"
			Case 3
				If page="main_style"  And i=0 Then 
					Response.Write "<input type=hidden name=""TemplateStr"" value="""
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write """>"
					Response.Write "���ֶ����ڻ������ã�  <a href=""?action=editmain&stype=2&StyleID="&StyleID&""">�������޸Ļ�������</a>"		
				Else
					
					Response.Write "<textarea name=""TemplateStr"" id=""t"&i&""" cols=""100"" rows=""5"">"
					Response.Write server.htmlencode(TemplateStr(i))
					Response.Write "</textarea>"
					Response.Write "</td><td class=""forumRowHighlight""><a href=""javascript:admin_Size(-5,'t"&i&"')""><img src=""images/minus.gif"" unselectable=""on"" border='0'></a> <a href=""javascript:admin_Size(5,'t"&i&"')""><img src=""images/plus.gif"" unselectable=""on"" border='0'></a><img src=images/viewpic.gif onclick=runscript(t"&i&")>"
				End If 
			End Select
		Response.Write "</td></tr>"
		
	Next
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=""3"">"
	Response.Write "</td></tr>"
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""�� ��"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=""2"">"
	Response.Write "<input type=""submit"" name=""B1"" value=""�� ��"">"
	Response.Write "</td></tr>"
	Response.Write "</table><p></p>"
	Response.Write "</form>"
	
	Rs.Close
	Set Rs=Nothing
End Sub
'�����̨ģ��
Sub Save_Main()
	Dim Page
	Dim TempStr,TemplateStr,stype
	stype=Dvbbs.checkStr(request("stype"))
	page=Dvbbs.checkStr(request("page"))
	If Not IsNumeric(stype) Then 
		Errmsg=ErrMsg + "<br /><li>�������ʽ����"
		Dvbbs_error()
	End If
	If page<>"page_admin" Then
		Errmsg=ErrMsg + "<br /><li>Ҫ�༭��ҳ��ģ���ֶ���δ������"
		Dvbbs_error()
	End If
	'��ȡ���е�����
	TemplateStr=""
	For Each TempStr in Request.form("TemplateStr")
		If TempStr<>"" Then 
			TemplateStr=TemplateStr&Replace(TempStr,"|||","")&"|||"
		End If
	Next 
	TemplateStr=Dvbbs.checkStr(Replace(TemplateStr,"@@@",""))
	If Trim(TemplateStr)="" Then 
		TemplateStr="|||"
	End If
	Set Rs=Dvbbs.Execute("Select H_ID,H_title,H_content From [Dv_help] Where H_ID=1")
	TempStr=Split(Dvbbs.checkStr(Rs(2)),"@@@")
	Select Case stype
		Case 1
			TemplateStr=TempStr(0)&"@@@"&TemplateStr&"@@@"&TempStr(2)
		Case 2
			TemplateStr=TempStr(0)&"@@@"&TempStr(1)&"@@@"&TemplateStr
		Case 3
			TemplateStr=TemplateStr&"@@@"&TempStr(1)&"@@@"&TempStr(2)
	End Select
	Set Rs=Nothing 
	Dvbbs.Execute("update [Dv_help] set H_content='"&TemplateStr&"' Where H_ID=1")
	Select Case stype
		Case 1
			Dv_suc(page&"���԰��޸ĳɹ�!")
		Case 2
			Dv_suc(page&"ͼƬ��Դ�޸ĳɹ�!")
		Case 3
			Dv_suc(page&"�������޸ĳɹ�!")
	End Select
End Sub
Function IsTruePage(page)
	IsTruePage=False
	If page<>"" Then 
		page=LCase(Trim(page))
		Dim myrs
		Set Myrs=Dvbbs.Execute("Select top 1 * From [Dv_Style]")
		For i= 2 to MyRs.Fields.Count-1
			If LCase( myrs(i).name)=page Then
				 IsTruePage=True
				 Exit For
			End If
		Next
	End If
End Function

Sub DelStyle()
	'����Ƿ��а���ʹ�ñ�ģ��
	If StyleID=SkinID Then 
		Errmsg=ErrMsg + "<br /><li>��ģ����Ĭ��ģ�棬������ɾ����"
		Dvbbs_error()
	End If
	Dim CssStyle,CssSid,Node
	Set Rs=Dvbbs.Execute("Select Forum_Css From Dv_setup")
	Set CssStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	CssStyle.Loadxml Rs(0)
	Dvbbs.Execute("Delete From [Dv_Style] Where ID="&StyleID&"")
	Dv_suc("�ɹ�ɾ����һ��ģ�塣")
	For Each Node in CssStyle.documentElement.selectNodes("css[tid="& StyleID&"]")
		CssStyle.documentElement.removeChild(node)
	Next
	Dim i
	i=1
	For Each Node in CssStyle.documentElement.selectNodes("css/@filename")
			Node.text="aspsky_"&i
			i=i+1
	Next
		i=1
		For Each Node in CssStyle.documentElement.selectNodes("css/@id")
			Node.text=i
			i=i+1
		Next
		Dvbbs.Execute("Update Dv_Setup Set Forum_Css='"&Dvbbs.Checkstr(CssStyle.xml)&"'")
	Dvbbs.loadSetup()
	Dvbbs.Loadstyle()
	createsccfile()
End Sub

Sub delpage()
	Dim StylePageName
	StylePageName=Dvbbs.checkStr(request("StylePageName"))
	If StylePageName="" Then 
		Errmsg=ErrMsg + "<br /><li>����д�ֶ�����"
		Dvbbs_error()
	End If
	If Not IsTruePage(StylePageName) Then 
		Errmsg=ErrMsg + "<br /><li>Ҫɾ�����ֶβ����ڡ�"
		Dvbbs_error()
	End If
	If IsSqlDataBase = 1 Then
		Dim i,Fieldslist,Rs
		Set Rs=Dvbbs.Execute("select * from Dv_Style")
		Fieldslist="id"
		For i= 1 to Rs.Fields.Count-1
			If LCase(Rs(i).name)<> LCase (StylePageName) Then 
				Fieldslist=Fieldslist&","&Rs(i).name
			End If
		Next
		Set Rs=Nothing 
		'�����������ݵ���ʱ��	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Style")
		'ɾ��ԭ�б�
		Dvbbs.Execute("Drop table Dv_Style")
		'�ٰ���ʱ���е����ݸ��ƹ���.
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_Style From Dv_tempatble ")
		'ɾ����ʱ��
		Dvbbs.Execute("Drop table Dv_tempatble")
		'�����������ݵ���ʱ��	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Stylehelp")
		'ɾ��ԭ�б�
		Dvbbs.Execute("Drop table Dv_Stylehelp")
		'�ٰ���ʱ���е����ݸ��ƹ���.
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_Stylehelp From Dv_tempatble ")
		'ɾ����ʱ��
		Dvbbs.Execute("Drop table Dv_tempatble") 
	Else
		Dvbbs.Execute("Alter Table [Dv_Style] Drop ["&StylePageName&"]")
		Dvbbs.Execute("Alter Table [Dv_Stylehelp] Drop ["&StylePageName&"]")
	End If 
	Dv_suc("ҳ��ģ��ɾ���ɹ���")
	Dvbbs.Loadstyle()
End Sub 
Sub addpage()
	Dim StylePageName
	StylePageName=Dvbbs.checkStr(request("StylePageName"))
	If StylePageName="" Then 
		Errmsg=ErrMsg + "<br /><li>����д�ֶ�����"
		Dvbbs_error()
	End If
	StylePageName="Page_"&StylePageName
	If IsTruePage(StylePageName) Then 
		Errmsg=ErrMsg + "<br /><li>��Ҫ������ģ���ֶ��Ѿ����ڡ�"
		Dvbbs_error()
	End If
	Dvbbs.Execute("alter table [Dv_Stylehelp] add "&StylePageName&" text not Null default'|||@@@|||@@@|||@@@|||'")
	Dvbbs.Execute("Update [Dv_Stylehelp] Set "&StylePageName&"='|||@@@|||@@@|||@@@|||'")
	Dvbbs.Execute("alter table [Dv_Style] add "&StylePageName&" text not Null default'|||@@@|||@@@|||@@@|||'")
	Dvbbs.Execute("Update [Dv_Style] Set "&StylePageName&"='|||@@@|||@@@|||@@@|||'")
	Dv_suc("��ҳ��ģ�崴���ɹ�!")
	Dvbbs.Loadstyle()
End Sub

Sub addstyle()
	Dim stylename,sql
	stylename=Dvbbs.checkStr(Request("stylename"))
	If Trim(stylename)=""  Then 
		 Errmsg=ErrMsg + "<br /><li>������ģ�����ơ�"
		Dvbbs_error()
	End If
	Set Rs=Dvbbs.Execute("select * From [Dv_Style] where ID="&StyleID)
	Dim styleFields,stylevalues
	styleFields="StyleName"
	stylevalues="'"&stylename&"'"
	For i= 2 to Rs.Fields.Count-1
		styleFields=styleFields&","&Rs(i).Name
		stylevalues=stylevalues&",'"&Dvbbs.checkStr(Rs(i))&"'"
	Next
	Set rs=Nothing
	sql="insert into [Dv_Style]("&styleFields&")values("&stylevalues&")"
	Dvbbs.Execute(SQL)
	Dim CssStyle,CssSid,Node,Node1
	Set Rs=Dvbbs.Execute("Select Forum_Css From Dv_setup")
	Set CssStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	CssStyle.Loadxml Rs(0)
	Set Rs=Dvbbs.Execute("select Max(id) from Dv_style")
	For Each Node in CssStyle.documentElement.selectNodes("css[tid="& StyleID&"]")
		Set Node1=node.cloneNode(True)
		Node1.selectSingleNode("tid").text=Rs(0)
		CssStyle.documentElement.appendChild(node1)
	Next
	Dim i
	i=1
	For Each Node in CssStyle.documentElement.selectNodes("css/@filename")
			Node.text="aspsky_"&i
			i=i+1
	Next
		i=1
		For Each Node in CssStyle.documentElement.selectNodes("css/@id")
			Node.text=i
			i=i+1
		Next
	Dvbbs.Execute("Update Dv_Setup Set Forum_Css='"&Dvbbs.Checkstr(CssStyle.xml)&"'")
	Dv_suc("��ģ�崴���ɹ�!")
	Dvbbs.loadSetup()
	Dvbbs.Loadstyle()
	createsccfile()
End Sub
Sub pagerename()
	Dim oldpagename,newpagename
	oldpagename=Dvbbs.checkStr(request("oldpagename"))
	newpagename=Dvbbs.checkStr(request("newpagename"))
	If LCase(newpagename)=LCase(oldpagename) Then
		Errmsg=ErrMsg + "<br /><li>��û�и�������"
		Dvbbs_error()
	End If
	If newpagename="" Then
		Errmsg=ErrMsg + "<br /><li>�����Ʋ���Ϊ��"
		Dvbbs_error()
	End If
	If OLDpagename="" Then
		Errmsg=ErrMsg + "<br /><li>���ύ�������Ǵ����."
		Dvbbs_error()
	End If
	oldpagename="page_"&oldpagename
	newpagename="page_"&newpagename
	If Not IsTruePage(oldpagename) Then 
		Errmsg=ErrMsg + "<br /><li>Ҫ�������ֶβ����ڡ�"
		Dvbbs_error()
	End If
	If IsTruePage(newpagename) Then 
		Errmsg=ErrMsg + "<br /><li>�ֶ�����"&newpagename&"�Ѿ���ռ��."
		Dvbbs_error()
	End If
		Dim i,Fieldslist,Rs
		Set Rs=Dvbbs.Execute("select * from Dv_Style")
		Fieldslist="id"
		For i= 1 to Rs.Fields.Count-1
			If LCase(Rs(i).name)<> LCase (oldpagename) Then 
				Fieldslist=Fieldslist&","&Rs(i).name
			Else
				Fieldslist=Fieldslist&","&Rs(i).name&" as "& newpagename
			End If
		Next
		Set Rs=Nothing
		'�����������ݵ���ʱ��	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Style")
		'ɾ��ԭ�б�
		Dvbbs.Execute("Drop table Dv_Style")
		'�ٰ���ʱ���е����ݸ��ƹ���.
		Dvbbs.Execute("Select * into Dv_Style From Dv_tempatble ")
		'ɾ����ʱ��
		Dvbbs.Execute("Drop table Dv_tempatble")
		'�����������ݵ���ʱ��	
		Dvbbs.Execute("Select "&Fieldslist&" into Dv_tempatble From Dv_Stylehelp")
		'ɾ��ԭ�б�
		Dvbbs.Execute("Drop table Dv_Stylehelp")
		'�ٰ���ʱ���е����ݸ��ƹ���.
		Dvbbs.Execute("Select * into Dv_Stylehelp From Dv_tempatble ")
		'ɾ����ʱ��
		Dvbbs.Execute("Drop table Dv_tempatble") 
	Dv_suc("�ɹ��� "&oldpagename&" �ֶθ���Ϊ "&newpagename&"")
End Sub
Sub rename()
	Dim stylename
	stylename=Dvbbs.checkStr(Request("stylename"))
	If Trim(stylename)=""  Then 
		Errmsg=ErrMsg + "<br /><li>�޸������������µ�ģ�����ơ�"
		Dvbbs_error()
	End If
	Dvbbs.Execute("update [Dv_Style] set StyleName='"&StyleName&"' where id="&StyleID&"")
	Dv_suc("ģ�����޸ĳɹ�!")
		Dvbbs.loadSetup()
	Dvbbs.Loadstyle()
End Sub
Sub EditCss()
	Dim Rs,CSSDOM,Css,tp
	Set Rs=Dvbbs.Execute("Select Forum_CSS From [Dv_Setup]")
	Set CSSDOM=Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
	CSSDOM.loadxml Rs(0)&""
	Rs.close
	Response.Write "<br /><table width=""95%"" border=""0"" cellspacing=""1"" cellpadding=""3"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>"
	Response.Write "��̳�����ʽ����"
	Response.Write "</th></tr>"
	Response.Write "<tr><td>"
	Response.Write "˵��:ͼƬ��·�����������""/"",�����ļ�����֧��д�ļ���������,�粻֧�������"
	Response.Write "</td></tr></table>"
	%>
	<SCRIPT LANGUAGE="JavaScript">
	<!--
	function EditCss(n)
	{
	n=n-1
		var trid=document.getElementsByName('CssTR');
		 for (var i=0;i<trid.length;i++)    {
			if (i!=n){
			trid[i].style.display="none";
			}
		   }
		EditTextarea.style.display = '';
		document.cssform.CssContent.value = document.cssform.CssBody[n].value;
		document.cssform.TempID.value = n;
		document.cssform.CssEdit[n].disabled=true;
	}
	
	function DllData(n){
		n=n-1
		if (document.cssform.DelCss[n].checked==true){
		document.cssform.CssName[n].value = '�ƻ���ɾ��...';
		document.cssform.CssName[n].disabled = true;
		document.cssform.CssBody[n].disabled = true;
		document.cssform.CssPicUrl[n].disabled = true;
		document.cssform.CssEdit[n].disabled=true;
		document.cssform.filename[n].disabled=true;
		}else{
		location.reload();
		}
	}
	
	function SubmitData(){
		var NewData,UpObject
		var e = document.cssform;
		NewData=e.CssContent.value;
		UpObject=e.TempID.value;
		if (NewData!=''){
			e.CssBody[UpObject].value=NewData;
		}
		for (var i=0;i<e.CssName.length;i++){
			if (e.CssName[i].value == '�ƻ���ɾ��...' || e.CssName[i].value == ''){
			e.CssName[i].value = '';
			e.CssBody[i].value = '';
			e.CssPicUrl[i].value = '';
			e.filename[i].value = '';
			}
		}
	}
	//-->
	</SCRIPT>
	<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder" >
	<form action="?action=savecss&StyleID=<%=StyleID%>" method="post" name="cssform" onsubmit="SubmitData();">
	<tr>
	<td width="5%" class="bodytitle" align=center>ID</td>
	<td width="15%" class="bodytitle" align=center>����</td>
	<td width="15%" class="bodytitle" align=center>��Ӧģ��</td>
	<td width="15%" class="bodytitle" align=center>ͼƬ��·��</td>
	<td width="15%" class="bodytitle" align=center>�����ļ�</td>
	<td width="15%" class="bodytitle" align=center>����</td>
	<td width="5%" class="bodytitle" align=center>ɾ��</td>
	</tr>
	<tr><td class="tableBorder1" colspan="6"></td></tr>
	<%
	For Each Css in CSSDOM.documentElement.selectNodes("css")
	%>
	<tr Name="CssTR" id="CssTR">
	<td class="forumRowHighlight" align="center"><%=Css.selectSingleNode("@id").text%></td>
	<td class="forumRow" align="center"><input type=text value="<%=Css.selectSingleNode("@type").text%>" name="CssName"></td>
	<td class="forumRow" align="center">
	<Select  Name="TemplateID" size="1">
	<%
	For Each tp in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<Option Value="""&tp.selectSingleNode("@id").text&""""
		If Not Css.selectSingleNode("tid[.='"& tp.selectSingleNode("@id").text &"']") Is Nothing Then
			Response.Write " selected "
		End If
		Response.Write ">"& tp.selectSingleNode("@stylename").text &"</Option>"
	Next
	%>
	</Select>
	</td>
	<td class="forumRowHighlight"  align="center">
	<%
	Response.Write "<input name=""CssPicUrl"" size=20 type=text value="""
	Response.Write css.selectSingleNode("@picurl").text
	Response.Write """>"
	%>
	</td>
	<td class="forumRow" align="center">
	<%
	Response.Write "<input name=""filename"" size=10 type=text value="""
	Response.Write Css.selectSingleNode("@filename").text
	Response.Write """>"%>.css
	</td>
	<td class="forumRow" align="center">
	<input type="button" value="�޸���ʽ����" name="CssEdit" onclick="EditCss(<%=Css.selectSingleNode("@id").text%>)">
	<div style="display:none">
	<textarea name="CssBody" id="CssBody" style="width:0;height=0" rows="0" ><%=Css.selectSingleNode("cssdata").text%></textarea>
	</div>
	</td>
	<td class="forumRowHighlight" align=center><INPUT TYPE="checkbox" NAME="DelCss" onclick="DllData(<%=Css.selectSingleNode("@id").text%>)" ></td>
	</tr>
	<%
	Next
	%>
	<tr>
	<td class="forumRowHighlight" align="center">�½�</td>
	<td class="forumRow" align="center"><input type=text value="" name="NewCssName"></td>
	<td class="forumRow" align="center">
	<Select  Name="NewTemplateID" size="1">
	<%
	i=0
	For Each tp in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
		Response.Write "<Option Value="""&tp.selectSingleNode("@id").text&""""
		If i=0 Then
			Response.Write " selected "
		End If
		Response.Write ">"& tp.selectSingleNode("@stylename").text &"</Option>"
		i=i+1
	Next
	%>
	</Select>
	</td>
	<td class="forumRowHighlight"  align="center">
	<%
	Response.Write "<input name=""newCssPicUrl"" size=20 type=text value="""
	Response.Write """>"
	%>
	</td>
	<td class="forumRow" align="center">
	<%
	Response.Write "<input name=""newfilename"" size=10 type=text value="""
	Response.Write """>"%>.css
	</td>
	<td class="forumRow" align="center" colspan="6">������д����,�����ٱ༭����.
	</td>
	</tr>
	<tr id="EditTextarea" style="display:none">
	<INPUT TYPE="hidden" NAME="TempID">
	<td height=400 class="forumRowHighlight" colspan=7>
	<textarea id="CssContent" style="width:100%" rows="30" ></textarea>
	</td>
	</tr>
	<%
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=7>"
	Response.Write "Css�ļ�Ŀ¼��"
	Response.Write "<input type=""text"" value=""skins/"" name=""cssfilepath"" readonly>"
	Response.Write "</td></tr>"
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"" colspan=7>"
	Response.Write "<input type=""submit"" name=""B1"" value=""�ύ�޸�"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "<input type=""button"" value=""�� ��"" onclick=""location.reload('template.asp?action=editcss&StyleID="&StyleID&"')"">"
	Response.Write "</td></tr>"
	%>
	</form></table>
	<%
	Set CSSDOM=Nothing 
End Sub

'//�����޸�CSSģ��
Sub savecss()
	If StyleID="" or not IsNumeric(StyleID) Then
		Errmsg=ErrMsg + "<br /><li>��ѡ����Ҫ�޸ĵ�CSS��ʽ!"
		Dvbbs_error()
		Exit Sub 
	End If
	Dim CssDom,RequestData,node,i,TemplateID,TemplateIDlist,filename,createfile
	Set CssDom=Server.CreateObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	Set Node=CssDom.appendChild(CssDom.createElement("xml"))
	Node.attributes.setNamedItem(CssDom.createNode(2,"cssfilepath","")).text=Request("cssfilepath")
	i=1
	For Each RequestData in Request.Form("CssName") 
		Set Node=CssDom.createNode(1,"css","")
		Node.attributes.setNamedItem(CssDom.createNode(2,"id","")).text=i
		Node.attributes.setNamedItem(CssDom.createNode(2,"type","")).text=RequestData
		filename=Request.Form("filename")(i)
		If InStr(filename,".") > 0 Then filename=""
		If filename<>"" Then createfile=True
		Node.attributes.setNamedItem(CssDom.createNode(2,"filename","")).text=filename
		Node.attributes.setNamedItem(CssDom.createNode(2,"picurl","")).text=Request.Form("CssPicUrl")(i)
		Node.appendChild(CssDom.createNode(1,"cssdata","")).text=Request.Form("CssBody")(i)
		Node.appendChild(CssDom.createNode(1,"tid","")).text=Request.Form("TemplateID")(i)	
		CssDom.documentElement.appendChild(node)
		i=i+1
	Next
	If Request.Form("NewCssName") <>"" Then
		Set Node=CssDom.createNode(1,"css","")
		Node.attributes.setNamedItem(CssDom.createNode(2,"id","")).text=i
		Node.attributes.setNamedItem(CssDom.createNode(2,"type","")).text=Request.Form("NewCssName")
		filename=Request.Form("newfilename")
		If InStr(filename,".") > 0 Then filename=""
		If filename<>"" Then createfile=True
		Node.attributes.setNamedItem(CssDom.createNode(2,"filename","")).text=filename
		Node.attributes.setNamedItem(CssDom.createNode(2,"picurl","")).text=Request.Form("newCssPicUrl")
		Node.appendChild(CssDom.createNode(1,"cssdata","")).text=""
		Node.appendChild(CssDom.createNode(1,"tid","")).text=Request.Form("NewTemplateID")
		CssDom.documentElement.appendChild(node)
	End If
	Dvbbs.Execute("Update [Dv_Setup] set Forum_CSS='"&Dvbbs.Checkstr(CssDom.xml)&"'")
	Dv_suc("��̳�����ʽ�޸ĳɹ���")
	Dvbbs.loadSetup()
	Dvbbs.Loadstyle()
	Set CSSDOM=Nothing 
	If createfile Then createsccfile()
End Sub
Sub createsccfile()
	On error resume Next
	Dim Fso,filename,Forum_CSS
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	If  Err Then
		err.Clear
		Errmsg=ErrMsg + "<br /><li>���ķ�������֧��д�ļ�,CSS�ļ�д��ʧ��,���ֹ�������������ļ����������!</li>"
		Dvbbs_error()
		Exit Sub
	End If
	For Each filename In Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css/@filename")
		If  filename.text<>"" Then
			If InStr(filename.text,".")=0 Then
				Dvbbs.SkinID=filename.selectSingleNode("../tid").text
				Dvbbs.LoadTemplates("")
				Forum_CSS=filename.selectSingleNode("../cssdata").text
				Forum_CSS=Replace(Forum_CSS,"{$width}",Dvbbs.mainsetting(0))
				Forum_CSS=Replace(Forum_CSS,"{$PicUrl}",filename.selectSingleNode("../@picurl").text)
				Fso.CreateTextFile(server.MapPath("../skins/"& filename.text &".css")).WriteLine(Forum_CSS)
				If  Err Then
					err.Clear
					Errmsg=ErrMsg + "<br /><li>���ķ�������֧��д�ļ�,CSS�ļ�д��ʧ��,���ֹ�������������ļ����������!</li>"
					Dvbbs_error()
					Exit Sub
				End If
			End If
		End If
	Next
	Set FSO=Nothing  
End Sub
Sub editmain()
	Dim stype,NowEditinfo
	Dim mystr
	stype=Request("stype")
	
	Select Case stype
		Case "1"
			NowEditinfo="���԰�"
			mystr="Dvbbs.lanstr"
		Case "2"
			NowEditinfo="��������"
			mystr="mainsetting"
		Case "3"
			NowEditinfo="HTTPͷ����"
			mystr="mainhtml(0)"
		Case "4"
			NowEditinfo="ҳ�濪ʼ����"
			mystr="mainhtml(1)"
		Case "5"
			NowEditinfo="����ͨ��"
			mystr="mainhtml(2)"
		Case "6"
			NowEditinfo="�������"
			mystr="mainhtml(3)"
		Case "7"
			NowEditinfo="������"
			mystr="mainhtml(5)"
		Case "8"
			NowEditinfo="��̳�˵�"
			mystr="mainhtml(6)"
		Case "9"
			mystr="mainhtml(4)"
			NowEditinfo="��������"
		Case "10"
			mystr="mainpic"
			NowEditinfo="ͼƬ����"
		Case Else
			Errmsg=ErrMsg + "<br /><li>���ύ�˴���Ĳ���."
			Dvbbs_error()	
	End Select
	Set Rs=Dvbbs.Execute("Select Main_Style ,StyleName From [Dv_Style] Where ID="&StyleID&"")
	Dim TemplateStr
	TemplateStr=Split(server.htmlencode(Rs(0)),"@@@")
	Response.Write "<form action=""?action=savemain&stype="&stype&"&StyleID="&StyleID&""" method=post>"
	Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=center class=""tableBorder"">"
	Response.Write "<tr>"
	Response.Write "<th width=""100%"" class=""tableHeaderText"" colspan=2 height=25>"
	Response.Write Rs(1)
	Response.Write NowEditinfo
	Response.Write "("&mystr&")����</th></tr>"
	Select Case stype
		Case "1"
			TemplateStr(0)=split(TemplateStr(0),"|||")
			For i=0 to UBound(TemplateStr(0))
				Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""left"" colspan=2>"
				Response.Write mystr&"("&i&")&nbsp;&nbsp;&nbsp;"		
				Response.Write "<input type=text name=""TemplateStr"" value="""&TemplateStr(0)(i)&""" Size=100>"
				Response.Write "</td></tr>"
			Next	
		Case "2"		
			TemplateStr(0)=split(TemplateStr(0),"|||")
			TemplateStr(0)(0)=split(TemplateStr(0)(0),"||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<table cellspacing=""1"" cellpadding=""0"" border=""0"" align=""left"" width=""100%"" >"
			Response.Write "<tr>"
			Response.Write "<td width=""300"" align=""right"" >����ȣ�</td>"
			Response.Write "<td align=""left"" class=""forumRowHighlight"" colspan=""3"" >"
			Response.Write "<input type=""text"" size=""5"" name=""TemplateStr"" value="""&TemplateStr(0)(0)(0)&""">&nbsp;(ʵ������:��<b>780px</b> һ��Ҫд��λ(px),���߰ٷֱ� ,��<b>98%</b>)&nbsp;"&mystr&"(0)"
			Response.Write "</td>"
			Response.Write "</tr>"
			Dim j,vtitle
			vtitle="aa|��������������ɫ��|��ʾ���ӵ�ʱ��������ӣ�ת�����ӣ��ظ��ȵ���ɫ��|��ҳ������ɫ��|һ���û�����������ɫ��|һ���û������ϵĹ�����ɫ��|��������������ɫ��|���������ϵĹ�����ɫ��|����Ա����������ɫ��|����Ա�����ϵĹ�����ɫ��|�������������ɫ��|��������ϵĹ�����ɫ��|���߿���ɫ��"
			vtitle=split(vtitle,"|")
			For j=1 to 12
				If j=4 or j=6 or j=8 or j=10 Then
					Response.Write "<input type=""hidden"" size=""10"" value="""&TemplateStr(0)(0)(j)&""" name=""TemplateStr"">"
				Else
					Response.Write "<tr>"
					Response.Write "<td height=""1"" colspan=""4"" bgcolor=""#6595D6""><td>"
					Response.Write "</tr>"
					Response.Write "<tr>"
					Response.Write "<td height=""25"" width=""300"" align=""right"" >"&vtitle(j)&"</td>"
					Response.Write "<td width=""20"" bgcolor="""&TemplateStr(0)(0)(j)&"""></td>"
					Response.Write "<td width=""180"" align=""left"">"
					Response.Write "<input type=""text"" size=""10"" value="""&TemplateStr(0)(0)(j)&""" name=""TemplateStr"">"&mystr&"("&j&")"
					Response.Write "</td>"
					Response.Write "<td></td>"
					Response.Write "</tr>"
				End If
			Next
			Response.Write "</table>"		
			Response.Write "</td></tr>"
			
		Case "3"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(0)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "4"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(1)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "5"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(2)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "6"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(3)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "7"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(5)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "8"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			TemplateStr(2)(6)=split(TemplateStr(2)(6),"###")
			For i=0 to UBound(TemplateStr(2)(6))
				Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
				Select Case i
					Case 0
						Response.Write "�ѵ�¼�û��˵�����<br />"		
						Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
						Response.Write TemplateStr(2)(6)(i)
						Response.Write "</textarea>"
						Response.Write "</td></tr>"
					Case 1
						Response.Write "δ��¼�û��˵�����<br />"		
						Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
						Response.Write TemplateStr(2)(6)(i)
						Response.Write "</textarea>"
						Response.Write "</td></tr>"
					Case 2 
						Response.Write "��չ�˵�����<br />"		
						Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
						Response.Write TemplateStr(2)(6)(i)
						Response.Write "</textarea>"
						Response.Write "</td></tr>"
					Case 3
						Response.Write "�˵��ָ�ͼƬ���ã�&nbsp;&nbsp;"		
						Response.Write "<input type=""text"" size=""20"" value="""&TemplateStr(2)(6)(i)&""" name=""TemplateStr"">&nbsp;"
						Response.Write "</td></tr>"
				End Select 
			Next		
		Case "9"
			TemplateStr(2)=split(TemplateStr(2),"|||")
			Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""center"" colspan=2>"
			Response.Write "<textarea name=""TemplateStr"" cols=""130"" rows=""15"">"	
			Response.Write TemplateStr(2)(4)
			Response.Write "</textarea>"
			Response.Write "</td></tr>"
		Case "10"
			TemplateStr(3)=split(TemplateStr(3),"|||")
			For i=0 to UBound(TemplateStr(3))
				Response.Write "<tr><td class=""forumRowHighlight"" height=40 align=""left"" colspan=2>"
				Response.Write mystr&"("&i&")&nbsp;&nbsp;&nbsp;"		
				Response.Write "<input type=text name=""TemplateStr"" value="""&TemplateStr(3)(i)&""" Size=100>"
				Response.Write "</td></tr>"
			Next	
		Case Else
				
	End Select
	Response.Write "<tr><td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""reset"" name=""Submit"" value=""�� ��"">"
	Response.Write "</td>"
	Response.Write "<td class=""forumRowHighlight"" height=""25"" align=""center"">"
	Response.Write "<input type=""submit"" name=""B1"" value=""�� ��"">"
	Response.Write "</table>"
	Response.Write "</form>"
	Set Rs=Nothing 
End Sub 
Sub savemain()
	Dim stype,NowEditinfo,TemplateStr,tempstr,Main_Style
	stype=Request("stype")
	TemplateStr=""
	Set Rs=Dvbbs.Execute("Select Main_Style From [Dv_Style] Where ID="&StyleID&"")
	Main_Style=Rs(0)
	Set rs=Nothing 
	Main_Style=split(Main_Style,"@@@")
	Select Case stype
		Case "1"
			NowEditinfo="���԰�"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then 
					TemplateStr=TemplateStr&TempStr&"|||"
				End If
			Next 
			TemplateStr=TemplateStr&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "2"
			NowEditinfo="��������"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then
					TemplateStr=TemplateStr&TempStr&"||"
				Else
					TemplateStr=TemplateStr&Chr(1)&"||"
				End If
			Next	
			TemplateStr=Left(TemplateStr,Len(TemplateStr)-2)
			Main_Style(0)=split(Main_Style(0),"|||")
			Dim i
			For i=1 to UBound(Main_Style(0))
				TemplateStr=TemplateStr&"|||"&Main_Style(0)(i)
			Next		
			TemplateStr=TemplateStr&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)'&"@@@"&Main_Style(3)
		Case "3"
			NowEditinfo="HTTPͷ����"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&TemplateStr&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "4"
			NowEditinfo="ҳ�濪ʼ����"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&TemplateStr&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "5"
			NowEditinfo="����ͨ��"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&TemplateStr&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
			Exit Sub 
		Case "6"
			NowEditinfo="�������"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&TemplateStr&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
		Case "7"
			NowEditinfo="������"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&TemplateStr&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
		Case "8"
			NowEditinfo="��̳�˵�"
			For Each TempStr in Request.form("TemplateStr")
				TemplateStr=TemplateStr&TempStr&"###"
			Next 
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&Main_Style(2)(4)&"|||"&Main_Style(2)(5)&"|||"&TemplateStr&"@@@"&Main_Style(3)
			Exit Sub 
		Case "9"
			NowEditinfo="��������"
			TemplateStr=Request.form("TemplateStr")
			Main_Style(2)=split(Main_Style(2),"|||")
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)(0)&"|||"&Main_Style(2)(1)&"|||"&Main_Style(2)(2)&"|||"&Main_Style(2)(3)&"|||"&TemplateStr&"|||"&Main_Style(2)(5)&"|||"&Main_Style(2)(6)&"@@@"&Main_Style(3)
		Case "10"
			NowEditinfo="ͼƬ����"
			For Each TempStr in Request.form("TemplateStr")
				If TempStr<>"" Then 
					TemplateStr=TemplateStr&TempStr&"|||"
				End If
			Next 
			TemplateStr=Main_Style(0)&"@@@"&Main_Style(1)&"@@@"&Main_Style(2)&"@@@"&TemplateStr
			Exit Sub 
		Case Else
			Errmsg=ErrMsg + "<br /><li>���ύ�˴���Ĳ���."
			Dvbbs_error()	
	End Select
	TemplateStr=Dvbbs.checkStr(TemplateStr)
	'Response.Write TemplateStr
	Dvbbs.Execute("Update [Dv_Style] set Main_Style='"&TemplateStr&"' Where ID="&StyleID&"")
	Dvbbs.Loadstyle()
	Dv_suc("��ģ��"&NowEditinfo&"�޸ĳɹ�!")
	If stype=2 Then
			createsccfile()
	End If
End Sub
'ģ���¡
Sub ghost()
	Dim oldStyleID,newStyleID
	oldStyleID=Request("oldStyleID")
	newStyleID=Request("newStyleID")
	If Not IsNumeric(newStyleID) or Not IsNumeric(oldStyleID) Then
		Errmsg=ErrMsg + "<br /><li>��������"
		Dvbbs_error()
		Exit Sub
	End If
	If newStyleID="" Or oldStyleID="" Then
		Errmsg=ErrMsg + "<br /><li>��������"
		Dvbbs_error()
		Exit Sub
	End If
	oldStyleID=CLng(oldStyleID)
	newStyleID=CLng(newStyleID)
	If newStyleID =	oldStyleID Then 
		Errmsg=ErrMsg + "<br /><li>Ŀ��ģ���Դģ�岻����ͬ��"
		Dvbbs_error()
		Exit Sub
	End If
	Dim styleFields,stylevalues,i,Sql
	Set Rs=Dvbbs.Execute("select * From [Dv_Style] where ID="&oldStyleID&"")
	If Rs.EOF Or Rs.BOF Then
		Errmsg=ErrMsg + "<br /><li>�޷�ȡ��Դģ������"
		Dvbbs_error()
		Exit Sub
	Else
		For i= 2 to Rs.Fields.Count-1
			styleFields=styleFields&Rs(i).Name
			stylevalues=stylevalues&"'"&Dvbbs.checkStr(Rs(i))&"'"
			Sql = Sql & Rs(i).Name &"= '"& Dvbbs.checkStr(Rs(i)) &"'"
			If i<Rs.Fields.Count-1 Then
				Sql=Sql&","
			End If
		Next
	End If
	Dvbbs.Execute("Update [Dv_Style] set "&Sql&" where ID="&newStyleID&" ")
	Dv_suc("ģ���¡�ɹ�!")
	Dim CssStyle,CssSid,Node,Node1
	Set Rs=Dvbbs.Execute("Select Forum_Css From Dv_setup")
	Set CssStyle=Server.CreateObject("Msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	CssStyle.Loadxml Rs(0)
	For Each Node in CssStyle.documentElement.selectNodes("css[tid="& newStyleID &"]")
		CssStyle.documentElement.removeChild(node)
	Next
	For Each Node in CssStyle.documentElement.selectNodes("css[tid="& oldStyleID &"]")
		Set Node1=node.cloneNode(True)
		Node1.selectSingleNode("tid").text=newStyleID
		CssStyle.documentElement.appendChild(node1)
	Next
	i=1
	For Each Node in CssStyle.documentElement.selectNodes("css/@filename")
			Node.text="aspsky_"&i
			i=i+1
	Next
		i=1
		For Each Node in CssStyle.documentElement.selectNodes("css/@id")
			Node.text=i
			i=i+1
		Next
	Dvbbs.Execute("Update Dv_Setup Set Forum_Css='"&Dvbbs.Checkstr(CssStyle.xml)&"'")
	Dvbbs.loadSetup()
	Dvbbs.Loadstyle()
	createsccfile()
End Sub
Sub RestoreBoardCache()
	Dim Board,node
	For Each node in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style/@id")
		Application.Contents.Remove(Dvbbs.CacheName & "_showtextads_"&node.text)
		For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text)
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text&"_-time")
		Next
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text)
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text&"_-time")
	Next
End Sub
%>