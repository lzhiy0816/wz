<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
	Head()
	dim admin_flag
	admin_flag=",4,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
		dvbbs_error()
	Else 
		select case request("action")
		case "save1"
			save1()
		case "save2"
			save2()
		case "save3"
			save3()
		case "search"
			search()
		case "view"
			view()
		case "edit"
			edit()
		case "del"
			del()
		case else
			consted()
		end select
		If founderr then call dvbbs_error()
		footer()
	End If 
sub consted()
dim sel
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height="23" colspan="3" class="tableHeaderText">��̳�������� | <a href="../boardhelp.asp" target=_blank><font color=#FFFFFF>Ԥ������</font></a></th>
</tr>
<FORM METHOD=POST ACTION="help.asp?action=search">
<tr>
<td height="25" colspan="3" class="forumrow">
<B>��������ؼ���</B>��<input type="text" name="keyword" size=50> <input type=submit name=submit value="�� ��"><BR><BR>��������ؼ��ֽ����������������ѯ����Ϊ��ʾ���а���
</td>
</tr>
</FORM>
</table><BR>
<%
search()
%><BR>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr><th height=23>��̳��������</th></tr>
<tr> 
<td width="48%" class="forumRowHighlight" valign=top>
<li>���������ݽ��Զ���ʾ��ǰ̨�İ���ҳ��;
<li>����һ�����෽����ѡȡ����ѡ�<font color=blue>��Ϊһ������</font>��������д�Ա߷�����⣬���������ݿ��Բ���д��
<li><font color=blue>������д���ݾ���ʹ��HTML�﷨��д��</font><BR>
</td>
</tr>
<tr> 
<td width="48%" class="forumRow" valign=top>
<FORM METHOD=POST ACTION="help.asp?action=save1">
<table width=100% >
<tr><td width=40>���⣺</td><td width=*><input type="text" name="title" size=50></td></tr>
<tr><td width=40>���ࣺ</td><td width=*>
<select name="classid">
<%
set rs=Dvbbs.Execute("select * from dv_help where h_type=0 and h_parentid=0")
do while not rs.eof
	Response.Write "<option value="&rs("h_id")&">"&server.htmlencode(rs("h_title"))&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
<option value="0">��Ϊһ������</option>
</select>&nbsp;<input type="text" name="classtitle" size=30>
</td></tr>
<tr><td width=40>���ͣ�</td><td width=*><input type=checkbox name="stype" checked value="1">&nbsp;ѡ��������������ݽ�����ʾ�ڰ����ļ���ҳ��</td></tr>
<tr><td width=40>���ݣ�</td><td width=*>
<textarea name="content" cols="80" rows="8" ID="TDcontent"></textarea><a href="javascript:admin_Size(-8,'TDcontent')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(8,'TDcontent')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</tr>
<tr><td width=40>&nbsp;</td><td width=*><input type=submit name=submit value="�� ��"></td></tr>
</table>
</FORM>
</td>
</tr>
</table>
<%
end sub

sub save1()
	dim title,content,parentid,stype
	if Request.form("classid")="0" then
		if request("classtitle")="" then
			Errmsg="�����ѡ������һ�����࣬����дѡ�����������������"
			founderr=true
			exit sub
		else
			title=request("classtitle")
		end if
		ParentID=0
	else
		if request("title")="" then
			Errmsg="����д��������"
			founderr=true
			exit sub
		else
			title=request("title")
		end if
		ParentID=Request.form("classid")
	end if
	if request("stype")="1" then
		stype=1
	else
		stype=0
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from dv_help"
	rs.open sql,conn,1,3
	rs.addnew
	rs("h_parentid")=parentid
	rs("h_title")=FilterJS(title)
	rs("h_content")=replace(FilterJS(request.form("content")),chr(10),"<br>")
	rs("h_bgimg")=FilterJS(request.form("targeturl"))
	rs("h_type")=0
	rs("h_stype")=stype
	rs.update
	rs.close
	set rs=nothing
	dv_suc("����ǰ̨�����ɹ���")
end sub

sub save3()
	dim title,content,parentid,stype
	if request("classid")="0" then
		if request("classtitle")="" then
			Errmsg="�����ѡ������һ�����࣬����дѡ�����������������"
			founderr=true
			exit sub
		else
			title=request("classtitle")
		end if
		ParentID=0
	else
		if request("title")="" then
			Errmsg="����д��������"
			founderr=true
			exit sub
		else
			title=request("title")
		end if
		ParentID=request("classid")
	end if
	if request("stype")="1" then
		stype=1
	else
		stype=0
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from dv_help where h_id="&request("id")
	rs.open sql,conn,1,3
	if not rs.eof then
	rs("h_parentid")=parentid
	rs("h_title")=FilterJS(title)
	rs("h_content")=replace(FilterJS(request("content")),chr(10),"<br>")
	rs("h_bgimg")=FilterJS(request("targeturl"))
	rs("h_type")=request("ctype")
	rs("h_stype")=stype
	end if
	rs.update
	rs.close
	set rs=nothing
	dv_suc("�����̨�����ɹ���")
end sub
Function FilterJS(v)
If not isnull(v) then
dim t
dim re
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern="(javascript)"
t=re.Replace(v,"&#106avascript")
re.Pattern="(jscript:)"
t=re.Replace(t,"&#106script:")
re.Pattern="(js:)"
t=re.Replace(t,"&#106s:")
re.Pattern="(value)"
t=re.Replace(t,"&#118alue")
re.Pattern="(about:)"
t=re.Replace(t,"about&#58")
re.Pattern="(file:)"
t=re.Replace(t,"file&#58")
re.Pattern="(document.cookie)"
t=re.Replace(t,"documents&#46cookie")
re.Pattern="(vbscript:)"
t=re.Replace(t,"&#118bscript:")
re.Pattern="(vbs:)"
t=re.Replace(t,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
t=re.Replace(t,"&#111n$2")
re.Pattern="(&#)"
t=re.Replace(t,"��#")
FilterJS=t
set re=nothing
End if
End Function

function search()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="5"  align=center class="tableBorder">

<tr> 
<th height="23" class="tableHeaderText" colspan=2>��̳�����ͺ�̨�˵������б�</th>
</tr>
<%
dim keyword,currentpage,page_count,totalrec
set rs=server.createobject("adodb.recordset")
if request("stype")<>"" then
	sql=" h_type="&request("stype")
end if
if request("keyword")<>"" then
	if sql<>"" then
	sql=sql & " and h_title like '%"&replace(request("keyword"),"'","")&"%' or h_content like '%"&replace(request("keyword"),"'","")&"%'"
	else
	sql=" h_title like '%"&replace(request("keyword"),"'","")&"%' or h_content like '%"&replace(request("keyword"),"'","")&"%'"
	end if
end if
if sql="" then
	sql="select * from dv_help where not h_id=1 order by h_id desc"
else
	sql="select * from dv_help where "&sql&" and not h_id=1 order by h_id desc"
end if
rs.open sql,conn,1,1

currentPage=request("page")
if currentpage="" or not IsNumeric(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if not rs.eof then
	rs.PageSize = 10
	rs.AbsolutePage=currentpage
	page_count=0
    totalrec=rs.recordcount
	while (not rs.eof) and (not page_count = rs.PageSize)
%>
<tr><td class=forumrowhighlight width=80% nowrap>
<B><%
Response.Write rs("h_title")
%></B>
</td>
<td class=forumrowhighlight width=20% align=right nowrap>
[<a href="help.asp?action=view&id=<%=rs(0)%>">�鿴����</a>] [<a href="help.asp?action=edit&id=<%=rs(0)%>">�༭</a>] [<a href="help.asp?action=del&id=<%=rs(0)%>">ɾ��</a>]
</td>
</tr>
<tr><td class=forumrow colspan=2>
<%
if not isnull(rs("h_content")) and rs("h_content")<>"" then
	Response.Write left(replace(replace(rs("h_content"),"<br>"," "),"<BR>"," "),100)
else
	if rs("H_ParentID")=0 then
	Response.Write "<font color=red>����Ϊһ���������!</font>"
	Else
	Response.Write "��������û��¼������!"
	End If
end if
%>
</td></tr>
<%
	page_count = page_count + 1
	rs.movenext
	wend
else
%>
<tr> 
<td height="23" class="forumrowhighlight" colspan=2>û���ҵ��κΰ���</td>
</tr>
<%
end if
rs.close
set rs=nothing
dim pcount,endpage
  	if totalrec mod 10=0 then
     		Pcount= totalrec \ 10
  	else
     		Pcount= totalrec \ 10+1
  	end if
	response.write "<tr><td valign=middle nowrap class=forumrowhighlight height=23 colspan=2>"
	response.write "<table width=""100%""><tr><td valign=middle nowrap class=forumrowhighlight>ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"
	response.write "&nbsp;ÿҳ<b>10</b> ����<b>"&totalrec&"</b></td>"
	response.write "<td valign=middle nowrap align=right class=forumrowhighlight>��ҳ��"
	if currentpage > 4 then
	response.write "<a href=""?page=1&action="&request("action")&"&keyword="&request("keyword")&"&stype="&request("stype")&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color=red>["&i&"]</font>"
		else
		response.write " <a href=""?page="&i&"&action="&request("action")&"&keyword="&request("keyword")&"&stype="&request("stype")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action="&request("action")&"&keyword="&request("keyword")&"&stype="&request("stype")&""">["&Pcount&"]</a>"
	end if
	response.write "</td></tr></table></td></tr>"
%>

</table>
<%
end function

function view()
%>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">

<tr> 
<th height="23" class="tableHeaderText">�鿴��̳����</th>
</tr><tr> 
<td height="23" class="forumrowhighlight">
<%
set rs=Dvbbs.Execute("select * from dv_help where h_id="&request("id"))
if rs.eof and rs.bof then
	Response.Write "û���ҵ�����"
else
	Response.Write "<BR><center><b>"&rs("h_title")&"</b></center><BR><BR><BR>"
	Response.Write "<blockquote>"&rs("h_content")&"</blockquote>"
	Response.Write "<div align=right>[<a href=help.asp?action=edit&id="&rs(0)&">�༭</a>] [<a href=help.asp?action=del&id="&rs(0)&">ɾ��</a>]</div>"
end if
%>
</td>
</tr>
</table>
<%
end function

function edit()
%>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">

<tr> 
<th height="23" class="tableHeaderText">��̳�����༭</th>
</tr><tr> 
<td height="23" class="forumrowhighlight">
<%
dim trs
set rs=Dvbbs.Execute("select * from dv_help where h_id="&request("id"))
if rs.eof and rs.bof then
	Response.Write "û���ҵ�����"
else
%>
<FORM METHOD=POST ACTION="help.asp?action=save3">
<input type=hidden value="<%=request("id")%>" name="id">
<table width=100% >
<tr><td width=40>���⣺</td><td width=*><input type="text" name="title" size=35 value="<%if not rs("h_parentid")=0 then Response.Write server.htmlencode(rs("h_title"))%>">&nbsp;&nbsp;
<select name="ctype" size=1>
<option value=0 <%if rs("h_type")=0 then Response.Write "selected"%>>ǰ̨</option>
<option value=1 <%if rs("h_type")=1 then Response.Write "selected"%>>��̨</option>
</select>
</td></tr>
<tr><td width=40>���ࣺ</td><td width=*>
<select name="classid">
<%
set trs=Dvbbs.Execute("select * from dv_help where h_parentid=0 order by H_type ")
do while not trs.eof
	Response.Write "<option value="&trs("h_id")
	if rs("h_parentid")=trs(0) then Response.Write " selected"
	Response.Write ">"&server.htmlencode(trs("h_title"))
	if Cint(trs("H_type")) = 1 Then Response.Write " [��̨]" Else Response.Write " [ǰ̨]" 
	Response.Write "</option>"
trs.movenext
loop
trs.close
set trs=nothing
%>
<option value="0" <%if rs("h_parentid")=0 then Response.Write "selected"%>>��Ϊһ������</option>
</select>&nbsp;<input type="text" name="classtitle" size=30 value="<%if rs("h_parentid")=0 then Response.Write server.htmlencode(rs("h_title"))%>">
</td></tr>
<tr><td width=40>���ͣ�</td><td width=*><input type=checkbox name="stype" value="1" <%if rs("h_stype")=0 then Response.Write "checked"%>>&nbsp;ѡ��������������ݽ�����ʾ����߲˵���(��̨��Ч)</td></tr>
<tr><td width=40>������</td><td width=*><input type="text" name="targeturl" size=35 value="<%=server.htmlencode(rs("h_bgimg"))%>"> ��д·��</td></tr>
<tr><td width=40>���ݣ�</td><td width=*>
<textarea name="content" cols="80" rows="8"><%if not isnull(rs("h_content")) and rs("h_content")<>"" then%><%=server.htmlencode(replace(replace(rs("h_content"),"<br>",chr(10)),"<BR>",chr(10)))%><%end if%></textarea>
</tr>
<tr><td width=40>&nbsp;</td><td width=*><input type=submit name=submit value="�����޸�"></td></tr>
</table>
</FORM>
<%
end if
%>
</td>
</tr>
</table>
<%
end function

function del()
	Dvbbs.Execute("delete from dv_help where (not h_id=1) and h_id="&Request("id"))
	dv_suc("ɾ����̳�����ɹ���")
end function
%>