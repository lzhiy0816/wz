<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
Head()
Dim admin_flag
Dim action
Dim sqlstr,l_type
l_type=request("l_type")
admin_flag=",3,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
	dvbbs_error()
Else 
If Request("action")="dellog" Then
	batch()
Else 
	Select Case l_type
	Case "3"
		sqlstr=" where l_type=3 "
		l_type=3
		main
	Case "4"
		sqlstr=" where l_type=4 "
		l_type=4
		main
	Case "5"
		sqlstr=" where l_type=5 "
		l_type=5
		main
	Case "6"
		sqlstr=" where l_type=6 "
		l_type=6
		main
	Case "0"
		sqlstr=" where l_type=0 "
		l_type=0
		main
	Case "1"
		sqlstr=" where l_type=1 "
		l_type=1
		main
	Case "2"
		sqlstr=" where l_type=2 "
		l_type=2
		main
	Case Else
		sqlstr=""
		l_type=""
		main
	End Select
End If
If founderr then call dvbbs_error()
footer()
End If 
Sub main()
Dim l_boardID
l_boardID=Request("l_boardID")
If l_boardID="" Then l_boardID="0"
If l_boardID<> 0 Then
	If sqlstr <> "" Then
		sqlstr=sqlstr &" and l_boardID="&l_boardID
	Else
		sqlstr=" where l_boardID="&l_boardID
	End If
End If
Dim keyword,checkvalue
checkvalue=Dvbbs.Checkstr(Request("checkvalue"))
keyword=Dvbbs.checkstr(Request("keyword"))
If keyword <> "" Then
	If checkvalue="" Then
		If sqlstr <> "" Then
			sqlstr=sqlstr &" and (l_touser like '%"&keyword&"%' Or l_content like '%"&keyword&"%' Or l_ip like '%"&keyword&"%' Or l_username like '%"&keyword&"%')"
		Else
			sqlstr=" where l_touser like '%"&keyword&"%' Or l_content like '%"&keyword&"%' Or l_ip like '%"&keyword&"%' Or l_username like '%"&keyword&"%'"
		End If
	Else
		If sqlstr <> "" Then
			sqlstr=sqlstr &" and "& checkvalue &" like '%"&keyword&"%'"
		Else
			sqlstr=" where "& checkvalue &" like '%"&keyword&"%'"
		End If
	End If
End If
%>
<form action=log.asp method=post">
<table cellPadding=1 cellSpacing=1 class=tableborder align=center>
<tr><th colspan=2 height=25>��̳��־�鿴</th></tr>
<tr><td class=tableHeaderText width="*">
&nbsp;�� �� �� Χ��<Select Name="l_boardid">
</Select>
&nbsp;��־���ͣ�<Select name="l_type">
<Option value="" 
<%If Request("l_type")="" Then Response.Write " selected"%>>ȫ����־</Option>
<Option value="3"<%If Request("l_type")="3" Then Response.Write " selected"%>>���ӹ���</Option>
<Option value="4"<%If Request("l_type")="4" Then Response.Write " selected"%>>�̶�����</Option>
<Option value="5"<%If Request("l_type")="5" Then Response.Write " selected"%>>���Ͳ���</Option>
<Option value="6"<%If Request("l_type")="6" Then Response.Write " selected"%>>�û�����</Option>
<Option value="0"<%If Request("l_type")="0" Then Response.Write " selected"%>>��̨��־A</Option>
<Option value="1"<%If Request("l_type")="1" Then Response.Write " selected"%>>��̨��־B</Option>
<Option value="2"<%If Request("l_type")="2" Then Response.Write " selected"%>>��̨��־C</Option>
</Select>
</td>
<td class=tableHeaderText width="100" align=Left>
</td>
</tr>
<tr><td class=forumrow width="*">&nbsp;&nbsp;&nbsp;&nbsp;�� �� �֣�<Input Type="text" Name="keyword" value="<%=Request("keyword")%>" Size="28">
&nbsp;�ؼ���ƥ�䣺<Select name="checkvalue">
<Option value="" <%If Request("checkvalue")="" Then Response.Write " selected"%>>ȫ����</Option>
<Option value="l_touser"<%If Request("checkvalue")="l_touser" Then Response.Write " selected"%>>��������</Option>
<Option value="l_content"<%If Request("checkvalue")="l_content" Then Response.Write " selected"%>>�¼�����</Option>
<Option value="l_ip"<%If Request("checkvalue")="l_ip" Then Response.Write " selected"%>>IP��ַ</Option>
<Option value="l_username"<%If Request("checkvalue")="l_username" Then Response.Write " selected"%>>������</Option>
</Select>
</td>
<td class=forumrow width="100" align=Left>
</td>
</tr>
<tr><td class=tableHeaderText width="100%" align=center><input type=submit value="�� ��">
</td>
</tr>
</table>
</form>
<SCRIPT LANGUAGE="JavaScript">
BoardJumpListSelect('<%=l_boardID%>',"l_boardid","������̳","",0)
</SCRIPT>
<br>
<%
Dim pagestr
Dim currentpage,page_count,Pcount,endpage
Dim sql,Rs,totalrec
currentPage=request("page")
If currentpage="" or not IsNumeric(currentpage) Then
	currentpage=1
Else
	currentpage=clng(currentpage)
End If
pagestr="?keyword="&Request("keyword")&"&l_type="& Request("l_type") &"&checkvalue="&Request("checkvalue") &"&l_boardID=" &Request("l_boardID")&"&"
Dvbbs.Forum_Setting(11)=50
sql="select * from [dv_log] "&sqlstr&" order by l_addtime desc"
'Response.Write SQL

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1

Response.Write "<table width=""95%"" border=""0"" cellspacing=""0"" cellpadding=""0""  align=center class=""tableBorder"" style=""word-break:break-all"" >"
Response.Write "<form action=log.asp?action=dellog&l_type="&l_type&" method=post name=even>"
Response.Write "<tr align=center>"
Response.Write "<th height=25 width=""10%"" >"
Response.Write "����"
Response.Write "</td>"
Response.Write "<th height=25 width=""45%"" >"
Response.Write "�¼�����"
Response.Write "</td>"
Response.Write "<th height=25 width=""15%"">"
Response.Write "����ʱ��"
Response.Write "</td>"
Response.Write "<th height=25 width=""15%"">"
Response.Write "IP"
Response.Write "</td>"
Response.Write "<th height=25 width=""10%"" >"
Response.Write "������"
Response.Write "</td>"
Response.Write "<th height=25 width=""5%"" >"
Response.Write "����"
Response.Write "</th>"
Response.Write "</tr>"
If Not(Rs.eof or Rs.bof) Then
	rs.PageSize = Dvbbs.Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
    	totalrec=rs.recordcount
	While (Not Rs.EOF) And (Not page_count = Rs.PageSize)
	Response.Write "<tr align=left>"
	Response.Write "<td class=""forumrow""  width=""10%"" >"
	Response.Write "<a href=../dispuser.asp?name="
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write " target=_blank>"
	Response.Write Dvbbs.HTMLEncode(rs("l_touser"))
	Response.Write "</a>"
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""45%"" >"
	Response.Write HighLigth(Dvbbs.HTMLEncode(URLDecode(Rs("l_content"))),keyword)
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""15%"">"
	Response.Write rs("l_addtime")
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""15%"">"
	Response.Write Rs("l_ip")
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""10%"">"
	Response.Write "<a href=../dispuser.asp?name="&Dvbbs.HTMLEncode(rs("l_username"))&" target=_blank>"&Dvbbs.HTMLEncode(rs("l_username"))&"</a>"
	Response.Write "</td>"
	Response.Write "<td class=""forumrow"" width=""5%"">"
	If Rs("l_type")<>2 Then
		Response.Write  "<input type=checkbox name=lid value="&rs("l_id")&">"
	End If
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td height=2></td></tr>"
	
	page_count = page_count + 1
	Rs.MoveNext
	Wend
	Response.Write "<tr><td class=forumrowHighLight colspan=6>��ѡ��Ҫɾ�����¼���<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=submit name=act value=ɾ��  onclick=""{if(confirm('��ȷ��ִ�еĲ�����?')){this.document.even.submit();return true;}return false;}"">"
	Response.Write "��<input type=submit name=act onclick=""{if(confirm('ȷ���������վ���еļ�¼��?')){this.document.even.submit();return true;}return false;}"" value=�����־></td></tr>"
	If totalrec mod Dvbbs.Forum_Setting(11)=0 Then
		Pcount= totalrec \ Dvbbs.Forum_Setting(11)
  	Else
  		Pcount= totalrec \ Dvbbs.Forum_Setting(11)+1
  	End If
  	Response.Write "<table border=0 cellpadding=0 cellspacing=3 width="""&Dvbbs.mainsetting(0)&""" align=center>"
  	Response.Write "<tr><td valign=middle nowrap>"
	Response.Write "ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"
	Response.Write "&nbsp;ÿҳ<b>"&Dvbbs.Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b></td>"
	Response.Write "<td valign=middle nowrap align=right>��ҳ��"
	If currentpage > 4 Then
		Response.Write "<a href="""&pagestr&"page=1"">[1]</a> ..."
	End If
	If Pcount>currentpage+3 Then
		endpage=currentpage+3
	Else
		endpage=Pcount
	End If
	For i=currentpage-3 to endpage
	If Not i<1 Then
		If i = clng(currentpage) Then
			response.write " <font color="&Dvbbs.mainsetting(1)&">["&i&"]</font>"
		Else
			Response.Write " <a href="""&pagestr&"page="&i&""">["&i&"]</a>"
		End If
	End If
	Next
	If currentpage+3 < Pcount Then   
		Response.Write "... <a href="""&pagestr&"page="&Pcount&""">["&Pcount&"]</a>"
	End If
	Response.Write "</td></tr></table>"
Else
	Response.Write "<tr align=center>"
	Response.Write "<td class=""forumrow"" width=""100%"" colspan=""6"" >"
	Response.Write "����ؼ�¼��"
	Response.Write "</td>"
	Response.Write "</tr>"
End If
Response.Write "</form>"
Response.Write "</table>"
Rs.close
Set rs=Nothing
End Sub

Sub batch()
	Dim lid
	If request("act")="ɾ��" Then
		If request.form("lid")="" Then
			DVbbs.AddErrmsg "��ָ������¼���"
		Else
			lid=replace(request.Form("lid"),"'","")
			lid=replace(lid,";","")
			lid=replace(lid,"--","")
			lid=replace(lid,")","")
		End If
	End if
	If request("act")="ɾ��" Then
		Dvbbs.Execute("delete from dv_log where Datediff(""D"",l_addtime, "&SqlNowString&") > 2 and l_id in ("&lid&")")
	ElseIf request("act")="�����־" Then
		If request("l_type")="" or IsNull(request("l_type")) Then 
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log Where Datediff(D,l_addtime, "&SqlNowString&") > 2")
			else
			Dvbbs.Execute("delete from dv_log Where Datediff('D',l_addtime, "&SqlNowString&") > 2")
			end if
		Else
			If IsSqlDataBase = 1 Then
			Dvbbs.Execute("delete from dv_log where  Datediff(D,l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			else
			Dvbbs.Execute("delete from dv_log where  Datediff('D',l_addtime, "&SqlNowString&") > 2 and l_type="&CInt(request("l_type"))&"")
			end if
		End If
	End If
	Dv_suc("�ɹ�ɾ����־��ע�⣺�����ڵ���־�ᱻϵͳ������")
End Sub

'�ؼ���ͻ����ʾ by ��ƮƮ
Function HighLigth(Str,keyword)
	If keyword="" Then
		HighLigth=Str
		Exit Function
	End IF
	Dim re
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	re.Pattern="("&keyword&")"
	HighLigth=re.Replace(Str,"<font color=""red"">$1</font>")
End Function

'URL���뺯�� by ��ƮƮ
Function URLDecode(enStr)
	On Error Resume Next
	Dim deStr,c,i,v:deStr=""
	For i=1 to len(enStr)
		c=Mid(enStr,i,1)
		If c="%" Then
			v=eval("&h"+Mid(enStr,i+1,2))
			If v<128 Then
				deStr=deStr&Chr(v)
				i=i+2
			Else
				If isvalidhex(Mid(enstr,i,3)) Then
					If isvalidhex(Mid(enstr,i+3,3)) Then	'����жϼ���Ƿ�˫�ֽ�--����
						v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
						deStr=deStr&Chr(v)
						i=i+5
					Else
						v=eval("&h"+Mid(enStr,i+1,2)+Cstr(Hex(Asc(Mid(enStr,i+3,1)))))	'--��
						deStr=deStr&Chr(v)
						i=i+3 
					End If
				Else 
					destr=destr&c
				End If
			End If
		Else
			If c="+" Then
				deStr=deStr&" "
			Else
				deStr=deStr&c
			End If
		End If
	Next
	URLDecode=deStr
End Function

Function IsValidHex(str)
	Dim c
	IsValidHex=True
	str=UCase(str)
	If Len(str)<>3 Then
		IsValidHex=False
		Exit Function
	End If
	If Left(str,1)<>"%" Then
		IsValidHex=False
		Exit Function
	End If
	c=Mid(str,2,1)
	If Not (((c>="0") And (c<="9")) Or ((c>="A") And (c<="Z"))) Then
		IsValidHex=False
		Exit Function
	End If
	c=Mid(str,3,1)
	If Not (((c>="0") And (c<="9")) Or ((c>="A") And (c<="Z"))) Then
		IsValidHex=False
		Exit Function
	End If
End Function
%>