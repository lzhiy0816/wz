<!--#include file="../conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	Head()
	dim aconn,aconnstr
	aConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(MyDbPath & "data/ipaddress.mdb")
	Set aconn = Server.CreateObject("ADODB.Connection")
	aconn.open aconnstr
	dim admin_flag
	admin_flag="29"
	If not Dvbbs.master or instr(","&session("flag")&",",",29,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й�����ҳ���Ȩ�ޡ�"
		dvbbs_error()
	Else
		Dim body
		dim dadadas
		Dim country,city,num
		dim ip1,ip2,oldip1,oldip2
		Dim esip,str1,sip,str2,str3,str4
		dim s1,s21,s2,s31,s3,s4
		set rs= server.createobject ("adodb.recordset")
		Call main()
		Footer()
	End if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=5 class="tableHeaderText">IP����Դ��������<a href="?action=add">���ӣɣе�ַ</a></th>
</tr>
<tr> 
<td width="20%" class=forumrow>ע������</td>
<td width="80%" class=forumrow colspan=4>�����Ҫ����IP������Դ��ֱ�����ӣ�������ӵ���Դ�����ݿ����Ѿ����ڣ�����ʾ���Ƿ�����޸ģ����ݿ�����û�еļ�¼��ֱ�����ӣ���Ҳ����ֱ�Ӷ����е����ݽ��й���������</td>
</tr>
<%
	if request("action") = "add" then 
		call addip()
	elseif request("action")="edit" then
		call editip()
	elseif request("action")="savenew" then
		call savenew()
	elseif request("action")="savedit" then
		call savedit()
	elseif request("action")="del" then
		call del()
	elseif request("action")="query" then
		call ipinfo()
	elseif request("action")="upip" then
		call upip()
	elseif request("action")="saveip" then
		call saveip()
	else
		call ipquery()
	end if
	response.write body
%>
</table>
<%
end sub

sub addip()
%>
<form action="?action=savenew" method=post>
<tr> 
    <th colspan="5" align=left>�޸� IP ��ַ</th>
</tr>
<tr> 
<td width="20%" class=forumrow>��ʼ I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip1" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>��β I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip2" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>��Դ����</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="country" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>��Դ����</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="city" size="30"></td>
</tr>
<tr> 
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow colspan=4><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</form>
<%
end sub

sub editip()
ip1=request("ip1")
ip2=request("ip2")
sql="select ip1,ip2,country,city from dv_address where ip1="&ip1&" and ip2="&ip2&""
rs.open sql,aconn,1,1
%>
<form action="?action=savedit" method=post>
<tr> 
    <th colspan="5" align=left>�޸� IP ��ַ</th>
</tr>
<tr> 
<td width="20%" class=forumrow>��ʼ I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip1" size="30" value="<%=deaddr(ip1)%>"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>��β I P</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="ip2" size="30" value="<%=deaddr(ip2)%>"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>��Դ����</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="country" size="30" value="<%=rs("country")%>"></td>
</tr>
<tr> 
<td width="20%" class=forumrow>��Դ����</td>
<td width="80%" class=forumrow colspan=4><input type="text" name="city" size="30" value="<%=rs("city")%>"></td>
</tr>
<input type="hidden" name="oldip1" value="<%=request("ip1")%>">
<input type="hidden" name="oldip2" value="<%=request("ip2")%>">
<tr> 
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow colspan=4><input type="submit" name="Submit" value="�� ��"></td>
</tr>
</form>
<%
end sub

sub savenew()
if request.form("ip1")="" then
body="<tr><td colspan=5 class=forumrow>����дIP��ַ��</td></tr>"
exit sub
end if
if request.form("ip2")="" then
body="<tr><td colspan=5 class=forumrow>����дIP��ַ��</td></tr>"
exit sub
end if
if request.form("country")="" or request.form("city")="" then
body="<tr><td colspan=5 class=forumrow>���Һͳ��б�����д��һ��</td></tr>"
exit sub
end if
ip1=enaddr(request.form("ip1"))
ip2=enaddr(request.form("ip2"))
country=request.form("country")
city=request.form("city")
sql="select ip1,ip2,country,city from dv_address where ip1<="&ip1&" and ip2>="&ip2&""
rs.open sql,aconn,1,3
if rs.eof and rs.bof then
rs.AddNew 
rs("ip1")=ip1
rs("ip2")=ip2
rs("country")=country
rs("city")=city
rs.update
body="<tr><td colspan=5 class=forumrow>���ӳɹ���<a href=?>�������������</a>��</td></tr>"
else
body="<tr><td colspan=5 class=forumrow>����ʧ�ܣ������Ѵ��ڣ��뵽������������ip��ַ�������޸ġ�</td></tr>"
end if
rs.close
end sub

sub savedit()
oldip1=int(request.form("oldip1"))
oldip2=int(request.form("oldip2"))
ip1=enaddr(request.form("ip1"))
ip2=enaddr(request.form("ip2"))
country=request.form("country")
city=request.form("city")
sql="select ip1,ip2,country,city from dv_address where ip1="&oldip1&" and ip2="&oldip2&""
rs.open sql,aconn,1,3
if not(rs.eof and rs.bof) then
rs("ip1")=ip1
rs("ip2")=ip2
rs("country")=country
rs("city")=city
rs.update
body="<tr><td colspan=5 class=forumrow>IP��ַ�޸ĳɹ���</td></tr>"
else
body="<tr><td colspan=5 class=forumrow>IP��ַ�޸�ʧ�ܣ�</td></tr>"
end if
rs.close
end sub

sub del()
ip1=request("ip1")
ip2=request("ip2")
sql="delete from dv_address where ip1="&ip1&" and ip2="&ip2&""
aconn.Execute(sql)
body="<tr><td colspan=5 class=forumrow>ɾ���ɹ���<a href=?>�������������</a>��</td></tr>"
end sub

sub ipinfo()
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not IsNumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	country=trim(request("country"))
	city=trim(request("city"))
	if request("ip")="" and country="" and city="" then
	sql="select ip1,ip2,country,city from dv_address order by ip1 desc,ip2 desc"
	else
	sql="select ip1,ip2,country,city from dv_address where "
	if request("ip")<>"" then
		num=enaddr(request("ip"))
		sql = sql&" ip1 <="&num&" and ip2 >="&num&" order by ip1 desc,ip2 desc"
	else
		if country<>"" then sql=sql&" country like '%"&country&"%'"
		if city<>"" then sql=sql&" city like '%"&city&"%'"
		sql=sql&" order by ip1 desc,ip2 desc"
	end if
	end if
	rs.open sql,aconn,1,1
	if rs.eof and rs.bof then
	call addip()
	else
%>
<tr> 
    <th colspan="5" align=left>IP ��ַ���������</th>
</tr>
  <tr align="center"> 
    <td width="20%" class="forumHeaderBackgroundAlternate"><B>��ʼ IP</B></td>
    <td width="20%" class="forumHeaderBackgroundAlternate"><B>��β IP</B></td>
    <td width="18%" class="forumHeaderBackgroundAlternate"><B>�� ��</B></td>
    <td width="30%" class="forumHeaderBackgroundAlternate"><B>�� ��</B></td>
    <td width="12%" class="forumHeaderBackgroundAlternate"><B>�� ��</B></td>
  </tr>
<%
		rs.PageSize = Cint(Dvbbs.Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Dvbbs.Forum_Setting(11)))
%>
    <tr>
    <td width="20%" class=forumrow><%=deaddr(rs("ip1"))%></td>
    <td width="20%" class=forumrow><%=deaddr(rs("ip2"))%></td>
    <td width="18%" class=forumrow><%=rs("country")%></td>
    <td width="30%" class=forumrow><%=rs("city")%></td>
    <td width="12%" align="center" class=forumrow><a href="?action=edit&ip1=<%=rs("ip1")%>&ip2=<%=rs("ip2")%>">�༭</a>|<a href="?action=del&ip1=<%=rs("ip1")%>&ip2=<%=rs("ip2")%>">ɾ��</a></td>
    </tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
%>
<tr><td colspan=5 class=forumrow align=center>��ҳ��
<%Pcount=rs.PageCount
	if currentpage > 4 then
	response.write "<a href=""?page=1&action=query&country="&country&"&city="&city&"&ip="&request("ip")&""">[1]</a> ..."
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
        response.write " <a href=""?page="&i&"&action=query&country="&country&"&city="&city&"&ip="&request("ip")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action=query&country="&country&"&city="&city&"&ip="&request("ip")&""">["&Pcount&"]</a>"
	end if
%>
</td>
</tr>
<%
	end if
	rs.close
end sub

sub ipquery()
%>
<form action="?action=query" method = post>
<tr> 
    <th colspan="5" align=left>IP ��ַ������</th>
</tr>
  <tr> 
    <td width="20%" class=forumrow>IP ��ַ��</td>
    <td width="80%" class=forumrow colspan=4> 
      <input type="text" name="ip" size=40>
    </td>
  </tr>
  <tr> 
    <td width="20%" class=forumrow>�� �ң�</td>
    <td width="80%" class=forumrow colspan=4> 
      <input type="text" name="country" size=40>
    </td>
  </tr>
  <tr> 
    <td width="20%" class=forumrow>�� �У�</td>
    <td width="80%" class=forumrow colspan=4> 
      <input type="text" name="city" size=40>
    </td>
  </tr>
  <tr> 
    <td width="20%" class=forumrow></td>
    <td width="80%" class=forumrow colspan=4> 
       <input type="submit" name="Submit" value="�� ��" class="button">
    </td>
  </tr>
</form>

<%end sub

function enaddr(sip)
 esip=cstr(sip)
 str1=Left(sip,CInt(InStr(sip,".")-1))
 sip=Mid(sip,cint(instr(sip,"."))+1)
 str2=Left(sip,cint(instr(sip,"."))-1)
 sip=mid(sip,cint(instr(sip,"."))+1)
 str3=left(sip,cint(instr(sip,"."))-1)
 str4=mid(sip,cint(instr(sip,"."))+1)
 enaddr=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
end function
function deaddr(sip)
sip=sip+1
s1=int(sip/256/256/256)
s21=s1*256*256*256
s2=int((sip-s21)/256/256)
s31=s2*256*256+s21
s3=int((sip-s31)/256)
s4=sip-s3*256-s31
deaddr=cstr(s1)+"."+cstr(s2)+"."+cstr(s3)+"."+cstr(s4)
end function
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>