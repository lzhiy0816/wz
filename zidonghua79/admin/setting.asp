<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<%	
	Head()
	dim admin_flag,rs_c
	admin_flag=",1,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then 
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		Call dvbbs_error()
	Else
		If request("action")="save" Then
			call saveconst()
		ElseIf request("action")="restore" Then
			Call restore()
		Else
			Call consted()
		end if
		Footer()
	end if

Sub consted()
Dim  sel
%>
<iframe width="260" height="165" id="colourPalette" src="../images/post/nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;border:1px gray solid" frameborder="0" scrolling="no" ></iframe>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<form method="POST" action="setting.asp?action=save" name="theform">
<tr> 
<th width="100%" colspan=3 class="tableHeaderText" height=25>��̳�������ã�Ŀǰֻ�ṩһ������)
</th></tr>
<tr> 
<td width="100%" class=Forumrow colspan=3 height=23>
<a href="#setting3">[������Ϣ]</a>&nbsp;<a href="#setting21">[��̳ϵͳ��������]</a>&nbsp;<a href="#setting6">[���Ļ�ѡ��]</a>&nbsp;<a href="#setting7">[��̳��ҳѡ��]</a>&nbsp;<a href="#setting8">[�û���ע��ѡ��]</a>&nbsp;<a href="#setting10">[ϵͳ����]</a>&nbsp;<a href="#setting12">[���ߺ��û���Դ]</a>
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=3 height=23>
<a href="#setting13">[�ʼ�ѡ��]</a>&nbsp;<a href="#setting14">[�ϴ�����]</a>&nbsp;<a href="#setting15">[�û�ѡ��(ǩ����ͷ�Ρ����е�)]</a>&nbsp;<a href="#setting16">[����ѡ��]</a>&nbsp;<a href="#setting17">[��ˢ�»���]</a>&nbsp;<a href="#setting18">[��̳��ҳ����]</a>&nbsp;<a href="#setting19">[��������]</a>
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=3 height=23>
<a href="#setting20">[����ѡ��]</a>&nbsp;<a href="#settingxu">[<font color=blue>�ٷ��������</font>]</a>&nbsp;<a href="#admin">[<font color=red>��ȫ����</font>]</a>&nbsp;<a href="challenge.asp">[<font color=blue>RSS/WAP/�ֻ�����/����֧��</font>]</a>
<a href="#SettingVIP">[VIP�û�������]</a>
</td>
</tr>
<tr> 
<td width="93%" class=bodytitle colspan=2 height=23>
���������̳�����ø����ˣ�����ʹ��<a href="?action=restore"><B>��ԭ��̳Ĭ������</B></a>
</td>
<input type="hidden" id="forum_return" value="<b>��ԭ��̳Ĭ������:</b><br><li>���������̳���ø����ˣ����Ե����ԭ��̳Ĭ�����ý��л�ԭ������<br><li>ʹ�ô˲�����ʹ��ԭ����������Ч����ԭ����̳��Ĭ�����ã���ȷ����������̳���ݻ��߼ǵû�ԭ�������Щ�������̳����Ҫ������">
<td class=bodytitle><a href=# onclick="helpscript(forum_return);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td width="50%" class=Forumrow>
<U>��̳Ĭ��ʹ�÷��</U></td>
<td width="43%" class=Forumrow>
<%
	Dim forum_sid,iforum_setting,stopreadme,forum_pack,iCssName,iCssID,iStyleName
	Dim Style_Option,css_Option,Forum_cid,TempOption
	set rs=dvbbs.execute("select forum_sid,forum_setting,forum_pack,forum_cid from dv_setup")
	Forum_sid=rs(0)
	Forum_pack=Split(rs(2),"|||")
	Iforum_setting=split(rs(1),"|||")
	Forum_cid=rs(3)
	Rs.close:Set Rs=Nothing
	stopreadme=iforum_setting(5)
%>
<Select Size=1 Name="sid">
<%
Dim Templateslist
For Each Templateslist in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_sid = CLng(Templateslist.selectSingleNode("@id").text) Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@stylename").text &" </Option>"
Next
%>
</Select>Css��ʽ:<Select Size=1 Name="cid">
<%
For Each Templateslist in Application(Dvbbs.CacheName & "_csslist").documentElement.selectNodes("css")
	Response.Write "<Option value="""& Templateslist.selectSingleNode("@id").text &""""
	If Forum_cid = CLng(Templateslist.selectSingleNode("@id").text)  Then Response.Write " selected "
	Response.Write ">"& Templateslist.selectSingleNode("@type").text &" </Option>"
Next
%>
</Select>
</td>
<input type="hidden" id="forum_skin" value="<b>��̳Ĭ��ʹ�÷��:</b><br><li>������������ѡ������̳��Ĭ��ʹ�÷��<br><li>�����ı���̳����뵽��̳���ģ������н����������">
<td class=Forumrow><a href=# onclick="helpscript(forum_skin);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=forumRowHighlight><U>��̳��ǰ״̬</U><BR>ά���ڼ�����ùر���̳</td>
<td class=forumRowHighlight> 
<input type=radio name="forum_setting(21)" value=0 <%if Dvbbs.forum_setting(21)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(21)" value=1 <%if Dvbbs.forum_setting(21)="1" then%>checked<%end if%>>�ر�&nbsp;
</td>
<input type="hidden" id="forum_open" value="<b>��̳��ǰ״̬:</b><br><li>�������Ҫ�����ĳ��򡢸������ݻ���ת��վ�����Ҫ��ʱ�ر���̳�Ĳ��������ڴ˴�ѡ��ر���̳��<br><li>�ر���̳�󣬿�ֱ��ʹ����̳��ַ��login.asp��¼��̳��Ȼ��ʹ����̳��ַ��admin_login.asp��¼��̨������д���̳�Ĳ���">
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_open);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow><U>ά��˵��</U><BR>����̳�ر��������ʾ��֧��html�﷨</td>
<td class=Forumrow> 
<textarea name="StopReadme" cols="50" rows="3" ID="TDStopReadme"><%=Stopreadme%></textarea><br><a href="javascript:admin_Size(-3,'TDStopReadme')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(3,'TDStopReadme')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
<input type="hidden" id="forum_opens" value="<b>��̳ά��˵��:</b><br><li>���������̳��ǰ״̬�йر�����̳�����ڴ�����ά��˵����������ʾ����̳��ǰ̨����Ա�������֪��̳�رյ�ԭ�����������ʹ��HTML�﷨��">
<td class=forumRow><a href=# onclick="helpscript(forum_opens);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=forumRowHighlight>
<U>��̳��ʱ���ã�</U></td>
<td class=forumRowHighlight> 
<input type=radio name="forum_setting(69)" value="0" <%If Dvbbs.forum_setting(69)="0" Then %>checked <%End If%>>�� ��</option>
<input type=radio name="forum_setting(69)" value="1" <%If Dvbbs.forum_setting(69)="1" Then %>checked <%End If%>>��ʱ�ر�
<input type=radio name="forum_setting(69)" value="2" <%If Dvbbs.forum_setting(69)="2" Then %>checked <%End If%>>��ʱֻ��
</td>
<input type="hidden" id="forum_isopentime" value="<b>��ʱ����ѡ��:</b><br><li>�����������������Ƿ����ö�ʱ�ĸ��ֹ��ܣ���������˱����ܣ������ú�����ѡ���е���̳����ʱ�䡣<br><li>����ڷǿ���ʱ������Ҫ���ı����ã���ֱ��ʹ����̳��ַ��login.asp��¼��̳��Ȼ��ʹ����̳��ַ��admin_login.asp��¼��̨������д���̳�Ĳ���">
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_isopentime);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow>
<U>��ʱ����</U><BR>�������Ҫѡ�񿪻��</td>
<td class=Forumrow> 
<%
Dvbbs.forum_setting(70)=split(Dvbbs.forum_setting(70),"|")
If UBound(Dvbbs.forum_setting(70))<2 Then 
	Dvbbs.forum_setting(70)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	Dvbbs.forum_setting(70)=split(Dvbbs.forum_setting(70),"|")
End If
For i= 0 to UBound(Dvbbs.forum_setting(70))
If i<10 Then Response.Write "&nbsp;"
%>
  <%=i%>�㣺<input type="checkbox" name="forum_setting(70)<%=i%>" value="1" <%If Dvbbs.forum_setting(70)(i)="1" Then %>checked<%End If%>>��
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="forum_opentime" value="<b>��̳����ʱ��:</b><br><li>���ñ�ѡ����ȷ�������˶�ʱ������̳���ܡ�<br><li>��������СʱΪ��λ������ذ��涨��ȷ��д<br><li>����ڷǿ���ʱ������Ҫ���ı����ã���ֱ��ʹ����̳��ַ��login.asp��¼��̳��Ȼ��ʹ����̳��ַ��admin_login.asp��¼��̨������д���̳�Ĳ���">
<td class=forumRow><a href=# onclick="helpscript(forum_opentime);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
</table><a name="admin"></a><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=3  align=Left id=tabletitlelink  height=25><b>��ȫ����</b>[<a href="#top">����</a>]
</th></tr>
<tr> 
<td class=forumRow width="50%">
<U>��̨����Ŀ¼���趨</U><br>ȱʡĿ¼ΪadminΪ��ȫ���,����������֪��Ŀ¼,���޸�.</td>
<td class=forumRow width="43%"> 
<input type="text" name="Forum_AdminFolder" size="35" value="<%=Dvbbs.CacheData(33,0)%>"><br><BR><b>ע��:</b>Ŀ¼���ƺ���Ҫ��"/",��"admin/"
<input type="hidden" id="AdminFolder" value="<b>��̨����Ŀ¼���趨:</b><br><li>��FTP���޸�������̳�Ĺ���,Ŀ¼���ơ�(ȱʡĿ¼Ϊadmin)<br><li>Ȼ�������޸Ĺ���Ŀ¼,����Ա��¼��̨��Ϳ����Զ������������趨��Ŀ¼.<br><li>������Ա��,�������޷�֪������ĵ�ַ.">
</td>
<td class=forumRowHighlight><a href=# onclick="helpscript(AdminFolder);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=forumRow width="50%">
<U>�Ƿ��ֹ�������������</U><BR>��ֹ��������������ܱ�������CC�����������ź�Ӱ��վ���������������ܵ����ԵĹ�����ʱ����</td>
<td class=forumRow width="43%"> 
<input type="radio" name="forum_setting(100)" value="0" <%if Dvbbs.forum_setting(100)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(100)" value="1" <%if Dvbbs.forum_setting(100)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="killcc" value="<b>�Ƿ��ֹ�������������:</b><br><li>��ֹ��������������ܱ�������CC�����������ź�Ӱ��վ���������������ܵ����ԵĹ�����ʱ������ƽʱ��رա�">
<td class=forumRow><a href=# onclick="helpscript(killcc);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=forumRow width="50%">
<U>����ͬһIP������Ϊ</U><BR>����ͬһIP�����������Լ��ٶ����CC������Ӱ�죬��������û����ʲ��㣬��������Ϊ0�رմ˹��ܣ����ܵ�������ʱ��ſ���</td>
<td class=forumRow width="43%"> 
<% Dim IP_MAX_value
If UBound(Dvbbs.forum_setting) > 101 Then
	IP_MAX_value=Dvbbs.forum_setting(101)
Else
	IP_MAX_value=0
End If
%>
<input type="text" name="forum_setting(101)" size="5" value="<%=IP_MAX_value%>">
</td>
<input type="hidden" id="IP_MAX" value="<b>����ͬһIP������:</b><br><li>����ͬһIP�����������Լ��ٶ����CC������Ӱ�죬��������û����ʲ��㣬��������Ϊ0�رմ˹��ܣ����ܵ�������ʱ��ſ��ţ�ƽʱ��رա�">
<td class=forumRow><a href=# onclick="helpscript(IP_MAX);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=3 class="tableHeaderText" height=25>�����ٷ��Զ�ͨѶ����
</th></tr>
<tr> 
<td class=forumRow width="50%">
<U>�Ƿ����ö����ٷ��Զ�ͨѶϵͳ</U><BR>�����������̳��̨�յ������ٷ�����֪ͨ�Լ�ֱ�Ӳ���ٷ����������ۺͷ���</td>
<td class=forumRow width="43%"> 
<input type=radio name="forum_pack(0)" value=0 <%if cint(forum_pack(0))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_pack(0)" value=1 <%if cint(forum_pack(0))=1 then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="forum_pack1" value="<b>�Ƿ����ö����Զ�����֪ͨϵͳ:</b><br><li>����������̨��������ʾ���������³��򡢲�����֪ͨ�ȡ�">
<td class=forumRow><a href=# onclick="helpscript(forum_pack1);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=ForumrowHighlight>
<U>����ͨѶϵͳ�û���������</U><BR>�û����������÷��š�|||���ֿ�<BR>�û����������뵽 <a href="http://bbs.dvbbs.net/Union_GetUserInfo.asp" target=_blank><font color=blue>�����ٷ�</font></a> ��ȡ</td>
<td class=ForumrowHighlight>
<%
If UBound(forum_pack)<2 Then ReDim forum_pack(3)
%>
<input type=text size=21 name="forum_pack(1)" value="<%=forum_pack(1)%>|||<%=forum_pack(2)%>">
</td>
<input type="hidden" id="forum_pack2" value="<b>����֪ͨϵͳ�û���������:</b><br><li>��Ҫ����֪ͨϵͳ�������ȵ������ٷ���̳ע��һ���û������ڶ����ٷ�֪ͨϵͳ��ȡ�����룬����д�ڴ������ɿ�����">
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_pack2);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>��̳������Ϣ</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_info(0)" size="35" value="<%=Dvbbs.Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��̳�ķ��ʵ�ַ</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(1)" size="35" value="<%=Dvbbs.Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��̳�Ĵ�������(��ʽ��YYYY-M-D)</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="forum_setting(74)" size="35" value="<%=Dvbbs.forum_setting(74)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>��̳��ҳ�ļ���</U></td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(11)" size="35" value="<%=Dvbbs.Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��վ��ҳ����</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(2)" size="35" value="<%=Dvbbs.Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>��վ��ҳ���ʵ�ַ</U></td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(3)" size="35" value="<%=Dvbbs.Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��̳����ԱEmail</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(5)" size="35" value="<%=Dvbbs.Forum_info(5)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>��ϵ���ǵ����ӣ�����дΪMailto����Ա��</U></td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(7)" size="35" value="<%=Dvbbs.Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��̳��ҳLogoͼƬ��ַ</U><BR>��ʾ����̳�������Ͻǣ��������·�����߾���·��</td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(6)" size="35" value="<%=Dvbbs.Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>վ��ؼ���</U><BR>������������������������վ�Ĺؼ�����<BR>ÿ���ؼ����á�|���ŷָ�</td>
<td width="50%" class=forumRow>  
<input type="text" name="Forum_info(8)" size="35" value="<%=Dvbbs.Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>վ������</U><BR>����������������˵������վ����Ҫ����<BR><font color=red>�������벻Ҫ��Ӣ�ĵĶ���</font></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_info(10)" size="35" value="<%=Dvbbs.Forum_info(10)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRow> <U>��̳��Ȩ��Ϣ</U></td>
<td width="50%" class=forumRow valign=top>  
<textarea name="Copyright" cols="50" rows="5" id=TdCopyright><%=Dvbbs.Forum_Copyright%></textarea>
<a href="javascript:admin_Size(-5,'TdCopyright')"><img src="images/minus.gif" unselectable="on" border='0'></a> <a href="javascript:admin_Size(5,'TdCopyright')"><img src="images/plus.gif" unselectable="on" border='0'></a>
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting21"></a><b>��̳ϵͳ��������</b>[<a href="#top">����</a>]--(������Ϣ�������û��޸�)</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳��Ա����</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_UserNum" size="25" value="<%=Dvbbs.CacheData(10,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��̳��������</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_TopicNum" size="25" value="<%=Dvbbs.CacheData(7,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳��������</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_PostNum" size="25" value="<%=Dvbbs.CacheData(8,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��̳����շ���</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_MaxPostNum" size="25" value="<%=Dvbbs.CacheData(12,0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����շ�������ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_MaxPostDate" size="25" value="<%=Dvbbs.CacheData(13,0)%>">(��ʽ��YYYY-M-D H:M:S)
</td>
</tr>
<tr> 
<td width="50%" class=forumRowHighlight> <U>��ʷ���ͬʱ���߼�¼����</U></td>
<td width="50%" class=forumRowHighlight>  
<input type="text" name="Forum_Maxonline" size="25" value="<%=Dvbbs.Maxonline%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��ʷ���ͬʱ���߼�¼����ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_MaxonlineDate" size="25" value="<%=Dvbbs.CacheData(6,0)%>">(��ʽ��YYYY-M-D H:M:S)
</td>
</tr>
</table><BR>

<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>���Ļ�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�¶���Ϣ��������</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(10)" value=0 <%if Dvbbs.forum_setting(10)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(10)" value=1 <%if Dvbbs.forum_setting(10)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����̳����Ϣ�Ƿ������֤��</U><BR>����������Է�ֹ�������Ϣ</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(80)" value=0 <%if Dvbbs.forum_setting(80)="0" Then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(80)" value=1 <%if Dvbbs.forum_setting(80)="1" Then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</table><BR>

<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting7"></a><b>��̳��ҳѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr>
<td width="50%" class=Forumrow>
<U>��ҳ��ʾ��̳���</U>
<input type="hidden" id="forum_depth" value="<b>��ҳ��ʾ��̳��Ȱ���:</b><br><li>0����һ����1����2�����Դ����ƣ�<li>���ù������̳��Ƚ�Ӱ����̳�������ܣ�������Լ���̳��������ã���������Ϊ1��">
</td>
<td width="43%" class=Forumrow> 
<input type=text size=10 name="forum_setting(5)" value="<%=Dvbbs.forum_setting(5)%>"> ��
</td>
<td class=Forumrow><a href=# onclick="helpscript(forum_depth);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=forumRowHighlight> <U>�Ƿ���ʾ�����ջ�Ա</U>
<input type="hidden" id="forum_userbirthday" value="<b>��ҳ��ʾ�����ջ�Ա����:</b><br><li>�������л�Ա����������ʾ����̳��ҳ��<li>���������ܽ�������Դ��">
</td>
<td class=forumRowHighlight>  
<input type=radio name="forum_setting(29)" value=0 <%if Dvbbs.forum_setting(29)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(29)" value=1 <%if Dvbbs.forum_setting(29)="1" then%>checked<%end if%>>��&nbsp;
</td>
<td class=forumRowHighlight><a href=# onclick="helpscript(forum_userbirthday);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
</table><BR>

<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>�û���ע��ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ��������û�ע��</U><BR>�رպ���̳������ע��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(37)" value=0 <%if Dvbbs.forum_setting(37)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if Dvbbs.forum_setting(37)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>ע���Ƿ������֤��</U><BR>����������Է�ֹ����ע��</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(78)" value=0 <%if Dvbbs.forum_setting(78)="0" Then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(78)" value=1 <%if Dvbbs.forum_setting(78)="1" Then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��¼�Ƿ������֤��</U><BR>����������Է�ֹ�����¼�½�����</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(79)" value=0 <%if Dvbbs.forum_setting(79)="0" Then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(79)" value=1 <%if Dvbbs.forum_setting(79)="1" Then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>��Աȡ�������Ƿ������֤��</U><BR>����������Է�ֹ�����¼�½�����</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(81)" value=0 <%if Dvbbs.forum_setting(81)="0" Then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(81)" value=1 <%if Dvbbs.forum_setting(81)="1" Then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��Աȡ�������������</U><BR>0���ʾ�����ƣ���ȡ���ʴ���󳬹������ƣ���ֹͣ��24Сʱ������ٴ�ʹ��ȡ�����빦�ܡ�</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(84)" size="3" value="<%=Dvbbs.forum_setting(84)%>">
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>����û�������</U><BR>��д���֣�����С��1����50</td>
<td width="50%" class=ForumRowHighlight>  
<input type="text" name="forum_setting(40)" size="3" value="<%=Dvbbs.forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��û�������</U><BR>��д���֣�����С��1����50</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(41)" size="3" value="<%=Dvbbs.forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>ͬһIPע����ʱ��</U><BR>�粻�����ƿ���д0</td>
<td width="50%" class=ForumRowHighlight>  
<input type="text" name="forum_setting(22)" size="3" value="<%=Dvbbs.forum_setting(22)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Email֪ͨ����</U><BR>ȷ������վ��֧�ַ���mail������������Ϊϵͳ�������</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(23)" value=0 <%if Dvbbs.forum_setting(23)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(23)" value=1 <%if Dvbbs.forum_setting(23)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>һ��Emailֻ��ע��һ���ʺ�</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(24)" value=0 <%if Dvbbs.forum_setting(24)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(24)" value=1 <%if Dvbbs.forum_setting(24)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>ע����Ҫ����Ա��֤</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(25)" value=0 <%if Dvbbs.forum_setting(25)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(25)" value=1 <%if Dvbbs.forum_setting(25)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>����ע����Ϣ�ʼ�</U><BR>��ȷ���������ʼ�����</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(47)" value=0 <%if Dvbbs.forum_setting(47)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(47)" value=1 <%if Dvbbs.forum_setting(47)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������Ż�ӭ��ע���û�</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(46)" value=0 <%if Dvbbs.forum_setting(46)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(46)" value=1 <%if Dvbbs.forum_setting(46)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>

</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>ϵͳ����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>��̳����ʱ��</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="Forum_info(9)" size="35" value="<%=Dvbbs.Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>������ʱ��</U></td>
<td width="50%" class=ForumRowHighlight>
<select name="forum_setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if i=CInt(Dvbbs.forum_setting(0)) then%>selected<%end if%>><%=i%></option>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�ű���ʱʱ��</U><BR>Ĭ��Ϊ300��һ�㲻������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(1)" size="3" value="<%=Dvbbs.forum_setting(1)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�Ƿ���ʾҳ��ִ��ʱ��</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(30)" value=0 <%If Dvbbs.forum_setting(30)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(30)" value=1 <%if Dvbbs.forum_setting(30)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow><U>��ֹ���ʼ���ַ</U><BR>������ָ�����ʼ���ַ������ֹע�ᣬÿ���ʼ���ַ�á�|�����ŷָ�<BR>������֧��ģ����������������eway��ֹ������ֹeway@aspsky.net����eway@dvbbs.net����������ע��</td>
<td width="50%" class=Forumrow> 
<input type="text" name="forum_setting(52)" size="50" value="<%=Dvbbs.forum_setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight><U>��̳�ű�������չ����</U><BR>������Ϊ����HTML���͵�ʱ��Խű������ʶ�����ã�<br>�����Ը�����Ҫ����Զ��Ĺ���<br>��ʽ�ǣ�������| �磺abc|efg| �����������abc��efg�Ĺ���</td>
<td width="50%" class=ForumRowHighlight> 
<Input type="text" name="forum_setting(77)" size="50" value="<%=Dvbbs.forum_setting(77)%>"><br> û����ӿ�����0,�����������һ���ַ�������"|"
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>���ߺ��û���Դ</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>������ʾ�û�IP</U><BR>�رպ���������û��顢��̳Ȩ�ޡ��û�Ȩ�����������û��������ɼ�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(28)" value=0 <%if Dvbbs.forum_setting(28)="0" then%>checked<%end if%>>����&nbsp;
<input type=radio name="forum_setting(28)" value=1 <%if Dvbbs.forum_setting(28)="1" then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>������ʾ�û���Դ</U><BR>�رպ���������û��顢��̳Ȩ�ޡ��û�Ȩ�����������û��������ɼ�<BR>���������ܽ�������Դ</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(36)" value=0 <%if Dvbbs.forum_setting(36)="0" then%>checked<%end if%>>����&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if Dvbbs.forum_setting(36)="1" then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������б���ʾ�û���ǰλ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(33)" value=0 <%if Dvbbs.forum_setting(33)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if Dvbbs.forum_setting(33)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>���������б���ʾ�û���¼�ͻʱ��</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(34)" value=0 <%if Dvbbs.forum_setting(34)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if Dvbbs.forum_setting(34)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������б���ʾ�û�������Ͳ���ϵͳ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(35)" value=0 <%If Dvbbs.forum_setting(35)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(35)" value=1 <%if Dvbbs.forum_setting(35)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>����������ʾ��������</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(15)" value=0 <%if Dvbbs.forum_setting(15)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(15)" value=1 <%if Dvbbs.forum_setting(15)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����������ʾ�û�����</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(14)" value=0 <%if Dvbbs.forum_setting(14)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(14)" value=1 <%if Dvbbs.forum_setting(14)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>ɾ������û�ʱ��</U><BR>������ɾ�����ٷ����ڲ���û�<BR>��λ�����ӣ�����������</td>
<td width="50%" class=ForumRowHighlight>  
<input type="text" name="forum_setting(8)" size="3" value="<%=Dvbbs.forum_setting(8)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>����̳����ͬʱ������</U><BR>�粻�����ƣ�������Ϊ0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(26)" size="6" value="<%=Dvbbs.forum_setting(26)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>չ���û������б�ÿҳ��ʾ�û���</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type="text" name="forum_setting(58)" size="6" value="<%=Dvbbs.forum_setting(58)%>">&nbsp;��
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting13"></a><b>�ʼ�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
	<td width="50%" class=Forumrow> <U>�����ʼ����</U>
	<input type="hidden" id="forum_emailplus" value="<b>�����ʼ��������:</b><br><li>ѡ�����ʱ��ȷ�Ϸ������Ƿ�֧�֡�">
	<BR>������ķ�������֧�������������ѡ��֧��</td>
	<td width="43%" class=Forumrow>  
	<select name="forum_setting(2)" id="forum_setting(2)" onChange="chkselect(options[selectedIndex].value,'know1');">
	<option value="0">��֧�� 
	<option value="1">JMAIL 
	<option value="2">CDONTS 
	<option value="3">ASPEMAIL 
	</select><div id=know1></div></td>
	<td class=Forumrow><a href=# onclick="helpscript(forum_emailplus);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
	<td class=ForumRowHighlight> <U>SMTP Server��ַ</U>
	<input type="hidden" id="forum_smtp" value="<b>SMTP Server��ַ����:</b><br><li>��ѡ�����ʼ����ʱ�������д�����磺smtp.21cn.com��<li>���ʼ���������ַ����д�Ǹ��ݹ���Ա����������������Ա����Ϊabc@163.net����������smtp.163.net��">
	<BR>ֻ������̳ʹ�������д��˷����ʼ����ܣ�����д���ݷ���Ч</td>
	<td class=ForumRowHighlight>  
	<input type="text" name="Forum_info(4)" size="35" value="<%=Dvbbs.Forum_info(4)%>">
	</td>
	<td class=forumRowHighlight><a href=# onclick="helpscript(forum_smtp);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
	<td class=Forumrow> <U>�ʼ���¼�û���</U><BR>ֻ������̳ʹ�������д��˷����ʼ����ܣ�����д���ݷ���Ч</td>
	<td colspan=2 class=Forumrow>
	<input type="text" name="Forum_info(12)" size="35" value="<%=Dvbbs.Forum_info(12)%>">
	</td>
</tr>
<tr> 
	<td class=ForumRowHighlight> <U>�ʼ���¼����</U></td>
	<td colspan=2 class=ForumRowHighlight>  
	<input type="password" name="Forum_info(13)" size="35" value="<%=Dvbbs.Forum_info(13)%>">
	</td>
</tr>
</table>
<a name="setting14"></a>
<BR>
<%
Dim UploadSetting
UploadSetting = Split(Dvbbs.forum_setting(7),"|")
If Ubound(UploadSetting)<>20 Then
	Redim UploadSetting(20)
	For i=0 to 20
		If i=2 or i=3 Then
			UploadSetting(i)=999
		Else
			UploadSetting(i)=0
		End If
	Next
End If
%>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><b>�ϴ�����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
	<td width="50%" class=Forumrow> <U>ͷ���ϴ�</U></td>
	<td width="43%" class=Forumrow>
	<SELECT name="UploadSetting(0)" id="UploadSetting(0)">
	<OPTION value=0 <%if UploadSetting(0)=0 then%>selected<%end if%>>��ȫ�ر�&nbsp;
	<OPTION value=1 <%if UploadSetting(0)=1 then%>selected<%end if%>>��ȫ��&nbsp;
	<OPTION value=2 <%if UploadSetting(0)=2 then%>selected<%end if%>>ֻ�����Ա�ϴ�&nbsp;
	</SELECT>
	</td>
	<input type="hidden" id="Forum_FaceUpload" value="<b>ͷ���ϴ�����:</b><br><li>�������˹��ܣ��û����԰�ͼ���ļ��ϴ�����������Ϊͷ��<li>���ϴ��������ж��ϴ�ͷ����й���<LI>��ȫ�رգ���ע����޸����϶��������ϴ�ͷ��<LI>��ȫ�򿪣���ע����޸����϶������ϴ�ͷ��<LI>ֻ�����Ա�ϴ�������Ա�޸ĸ�������ʱ�����ϴ�ͷ��">
	<td class=Forumrow><a href=# onclick="helpscript(Forum_FaceUpload);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>��������ͷ���ļ���С</U></td>
	<td class=ForumRowHighlight> 
	<input type="text" name="UploadSetting(1)" size="6" value="<%=UploadSetting(1)%>">&nbsp;K
	</td>
	<input type="hidden" id="Forum_FaceUploadSize" value="<b>ͷ���ļ���С����:</b><br><li>�����ϴ�ͷ���ļ��Ĵ�С��<li>�û�ͷ����ϴ������⣬��鿴���û�ѡ�������á�">
	<td class=ForumRowHighlight><a href=# onclick="helpscript(Forum_FaceUploadSize);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
	<td class=Forumrow ><U>ѡȡ�ϴ����:</U></td>
	<td class=Forumrow >
	<select name="UploadSetting(2)" id="UploadSetting(2)" onChange="chkselect(options[selectedIndex].value,'know2');">
	<option value="999">�ر�
	<option value="0">������ϴ���
	<option value="1">Aspupload3.0��� 
	<option value="2">SA-FileUp 4.0���
	<option value="3">DvFile-Up V1.0���
	</option></select><div id="know2"></div>
	</td>
	<td class=Forumrow >
	<input type="hidden" id="forum_upload" value="<b>ѡȡ�ϴ��������:</b><br><li>��ѡȡʱ����̳ϵͳ���Զ�Ϊ�����������Ƿ�֧�ָ������<li>����ʾ��֧�֣���ѡ��رա�">
	<a href=# onclick="helpscript(forum_upload);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>ѡȡ����Ԥ��ͼƬ���:</U></td>
	<td class=ForumRowHighlight> 
	<select name="UploadSetting(3)" id="UploadSetting(3)" onChange="chkselect(options[selectedIndex].value,'know3');">
	<option value="999">�ر�
	<option value="0">CreatePreviewImage���
	<option value="1">AspJpeg���
	<option value="2">SA-ImgWriter���
	</select><div id="know3"></div>
	</td>
	<td class=forumRowHighlight>
	<input type="hidden" id="forum_CreatImg" value="<b>ѡȡ����Ԥ��ͼƬ�������:</b><br><li>��ѡȡʱ����̳ϵͳ���Զ�Ϊ�����������Ƿ�֧�ָ������<li>����ʾ��֧�֣���ѡ��رա�">
	<a href=# onclick="helpscript(forum_CreatImg);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
	<td class=ForumRow><U>����Ԥ��ͼƬ��С����(���|�߶�):</U></td>
	<td class=ForumRow>
		��ȣ�<INPUT TYPE="text" NAME="UploadSetting(14)" size=10 value="<%=UploadSetting(14)%>"> ����
		�߶ȣ�<INPUT TYPE="text" NAME="UploadSetting(15)" size=10 value="<%=UploadSetting(15)%>"> ����
	</td>
	<td class=ForumRow>
	<input type="hidden" id="forum_CreatImgSize" value="<b>����Ԥ��ͼƬ��С���ð���:</b><br><li>��ѡȡ������Ԥ��ͼƬ��������ҷ�������װ����Ӧ������˹��ܲ�����Ч��">
	<a href=# onclick="helpscript(forum_CreatImgSize);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>����Ԥ��ͼƬ��С����ѡ��:</U></td>
	<td class=ForumRowHighlight> 
		<SELECT name="UploadSetting(16)" id="UploadSetting(16)">
		<OPTION value=0>�̶�</OPTION>
		<OPTION value=1>�ȱ�����С</OPTION>
		</SELECT>
	</td>
	<td class=ForumRowHighlight></td>
</tr>
<tr> 
	<td class=ForumRow><U>ͼƬˮӡ���ÿ���:</U></td>
	<td class=ForumRow> 
		<SELECT name="UploadSetting(17)" id="UploadSetting(17)">
		<OPTION value="0">�ر�ˮӡЧ��</OPTION>
		<OPTION value="1">ˮӡ����Ч��</OPTION>
		<OPTION value="2">ˮӡͼƬЧ��</OPTION>
		</SELECT>
	</td>
	<td class=ForumRow></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>�ϴ�ͼƬ���ˮӡ������Ϣ����Ϊ�ջ�0��:</U></td>
	<td class=ForumRowHighlight> 
	<INPUT TYPE="text" NAME="UploadSetting(4)" size=40 value="<%=UploadSetting(4)%>">
	</td>
	<td class=ForumRowHighlight>
	<input type="hidden" id="forum_CreatText" value="<b>�ϴ�ͼƬ���ˮӡ���ְ���:</b><br><li>������Ҫˮӡ����Ч����������Ϊ�գ�<li>ˮӡ�����������˳���15���ַ�,��֧���κ�WEB�����ǣ�<li>Ŀǰ֧�ֵ����ͼƬ����У�AspJpeg�����SA-ImgWriter V1.21���.">
	<a href=# onclick="helpscript(forum_CreatText);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
	<td class=ForumRow><U>�ϴ����ˮӡ�����С:</U></td>
	<td class=ForumRow> 
	<INPUT TYPE="text" NAME="UploadSetting(5)" size=10 value="<%=UploadSetting(5)%>"> <b>px</b>
	</td>
	<td class=ForumRow></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>�ϴ����ˮӡ������ɫ:</U></td>
	<td class=ForumRowHighlight>
	<input type="hidden" name="UploadSetting(6)" id="UploadSetting(6)" value="<%=UploadSetting(6)%>">
	<img border=0 src="../images/post/rect.gif" style="cursor:pointer;background-Color:<%=UploadSetting(6)%>;" onclick="Getcolor(this,'UploadSetting(6)');" title="ѡȡ��ɫ!">
	</td>
	<td class=ForumRowHighlight></td>
</tr>
<tr> 
	<td class=ForumRow><U>�ϴ����ˮӡ��������:</U></td>
	<td class=ForumRow>
	<SELECT name="UploadSetting(7)" id="UploadSetting(7)">
	<option value="����">����</option>
	<option value="����_GB2312">����</option>
	<option value="������">������</option>
	<option value="����">����</option>
	<option value="����">����</option>
	<OPTION value="Andale Mono" selected>Andale Mono</OPTION> 
	<OPTION value=Arial>Arial</OPTION> 
	<OPTION value="Arial Black">Arial Black</OPTION> 
	<OPTION value="Book Antiqua">Book Antiqua</OPTION>
	<OPTION value="Century Gothic">Century Gothic</OPTION> 
	<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
	<OPTION value="Courier New">Courier New</OPTION>
	<OPTION value=Georgia>Georgia</OPTION>
	<OPTION value=Impact>Impact</OPTION>
	<OPTION value=Tahoma>Tahoma</OPTION>
	<OPTION value="Times New Roman" >Times New Roman</OPTION>
	<OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
	<OPTION value="Script MT Bold">Script MT Bold</OPTION>
	<OPTION value=Stencil>Stencil</OPTION>
	<OPTION value=Verdana>Verdana</OPTION>
	<OPTION value="Lucida Console">Lucida Console</OPTION>
	</SELECT>
	</td>
	<td class=ForumRow></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>�ϴ�ˮӡ�����Ƿ����:</U></td>
	<td class=ForumRowHighlight> 
		<SELECT name="UploadSetting(8)" id="UploadSetting(8)">
		<OPTION value=0>��</OPTION>
		<OPTION value=1>��</OPTION>
		</SELECT>
	</td>
	<td class=ForumRowHighlight></td>
</tr>
<!-- �ϴ�ͼƬ���ˮӡLOGOͼƬ���� -->
<tr> 
	<td class=ForumRow><U>�ϴ�ͼƬ���ˮӡLOGOͼƬ��Ϣ����Ϊ�ջ�0��:</U><br>��дLOGO��ͼƬ���·��</td>
	<td class=ForumRow> 
	<INPUT TYPE="text" NAME="UploadSetting(9)" size=40 value="<%=UploadSetting(9)%>">
	</td>
	<td class=ForumRow></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>�ϴ�ͼƬ���ˮӡ͸����:</U></td>
	<td class=ForumRowHighlight> 
	<INPUT TYPE="text" NAME="UploadSetting(10)" size=10 value="<%=UploadSetting(10)%>"> ��60%����д0.6
	</td>
	<td class=ForumRowHighlight></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>ˮӡͼƬȥ����ɫ:</U><br>����Ϊ����ˮӡͼƬ��ȥ����ɫ��</td>
	<td class=ForumRowHighlight> 
	<INPUT TYPE="text" NAME="UploadSetting(18)" ID="UploadSetting(18)" size=10 value="<%=UploadSetting(18)%>"> 
	<img border=0 src="../images/post/rect.gif" style="cursor:pointer;background-Color:<%=UploadSetting(18)%>;" onclick="Getcolor(this,'UploadSetting(18)');" title="ѡȡ��ɫ!">
	</td>
	<td class=ForumRowHighlight></td>
</tr>
<tr> 
	<td class=ForumRow><U>ˮӡ���ֻ�ͼƬ�ĳ���������:</U><br>��ˮӡͼƬ�Ŀ�Ⱥ͸߶ȡ�</td>
	<td class=ForumRow> 
	��ȣ�<INPUT TYPE="text" NAME="UploadSetting(11)" size=10 value="<%=UploadSetting(11)%>"> ����
	�߶ȣ�<INPUT TYPE="text" NAME="UploadSetting(12)" size=10 value="<%=UploadSetting(12)%>"> ����
	</td>
	<td class=ForumRow></td>
</tr>
<tr> 
	<td class=ForumRowHighlight><U>�ϴ�ͼƬ���ˮӡLOGOλ������ :</U></td>
	<td class=ForumRowHighlight>
	<SELECT NAME="UploadSetting(13)" id="UploadSetting(13)">
		<option value="0">����</option>
		<option value="1">����</option>
		<option value="2">����</option>
		<option value="3">����</option>
		<option value="4">����</option>
	</SELECT>
	</td>
	<td class=ForumRowHighlight></td>
</tr>
<!-- �ϴ�ͼƬ���ˮӡLOGOͼƬ���� -->
<%
If IsObjInstalled("Scripting.FileSystemObject") Then 
%>
<tr> 
<td class=ForumRow><U>�Ƿ�����ļ���ͼƬ������</U></td>
<td class=ForumRow>
<input type=radio name="Forum_Setting(75)" value=0 <%if Dvbbs.Forum_Setting(75)=0 Then %>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(75)" value=1 <%if Dvbbs.Forum_Setting(75)=1 Then %>checked<%end if%>>��&nbsp;
</td>
<td class=ForumRow>
</td>
</tr>
<tr> 
<td class=ForumRowHighlight><U>�ϴ�Ŀ¼�趨</U></td>
<td class=ForumRowHighlight>
<%
If Dvbbs.Forum_Setting(76)="" Or Dvbbs.Forum_Setting(76)="0" Then Dvbbs.Forum_Setting(76)="UploadFile/"
%>
<input type=text name="Forum_Setting(76)" value=<%=Dvbbs.Forum_Setting(76)%>>����޸��˴������FTP�ֹ�����Ŀ¼���ƶ�ԭ���ϴ��ļ���
</td>
<td class=ForumRowHighlight>
</td>
</tr>
<%
End If 
%>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting15"></a><b>�û�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�������ǩ��</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(42)" value=0 <%if Dvbbs.forum_setting(42)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(42)" value=1 <%if Dvbbs.forum_setting(42)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�����û�ʹ��ͷ��</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(53)" value=0 <%if Dvbbs.forum_setting(53)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(53)" value=1 <%if Dvbbs.forum_setting(53)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>���ͷ��ߴ�</U><BR>��������Ϊͷ������߶ȺͿ��</td>
<td class=Forumrow width="50%"> 
<input type="text" name="forum_setting(57)" size="6" value="<%=Dvbbs.forum_setting(57)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>Ĭ��ͷ����</U><BR>��������Ϊ��̳ͷ���Ĭ�Ͽ��</td>
<td width="50%" class=ForumRowHighlight>  
<input type="text" name="forum_setting(38)" size="6" value="<%=Dvbbs.forum_setting(38)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>Ĭ��ͷ��߶�</U><BR>��������Ϊ��̳ͷ���Ĭ�Ͽ��</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(39)" size="6" value="<%=Dvbbs.forum_setting(39)%>">&nbsp;����
</td>
</tr>
<tr> 
<td class=ForumRowHighlight width="50%"><U>ʹ���Զ���ͷ������ٷ�����</U></td>
<td class=ForumRowHighlight width="50%"> 
<input type="text" name="forum_setting(54)" size="6" value="<%=Dvbbs.forum_setting(54)%>">&nbsp;ƪ
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���������վ������ͷ��</U><BR>�����Ƿ����ֱ��ʹ��http..������url��ֱ����ʾͷ��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(55)" value=0 <%if Dvbbs.forum_setting(55)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(55)" value=1 <%if Dvbbs.forum_setting(55)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�û�ǩ���Ƿ���UBB����</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(65)" value=0 <%if Dvbbs.forum_setting(65)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(65)" value=1 <%if Dvbbs.forum_setting(65)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ǩ���Ƿ���HTML����</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(66)" value=0 <%if Dvbbs.forum_setting(66)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(66)" value=1 <%if Dvbbs.forum_setting(66)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�û��Ƿ�����ͼ��ǩ</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(67)" value=0 <%if Dvbbs.forum_setting(67)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(67)" value=1 <%if Dvbbs.forum_setting(67)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û��Ƿ���Flash��ǩ</U></td>
<td width="50%" class=Forumrow>  
<input type=radio name="forum_setting(71)" value=0 <%if Dvbbs.forum_setting(71)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(71)" value=1 <%if Dvbbs.forum_setting(71)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�û�ͷ��</U><BR>�Ƿ������û��Զ���ͷ��</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(6)" value=0 <%if Dvbbs.forum_setting(6)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(6)" value=1 <%if Dvbbs.forum_setting(6)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�û�ͷ����󳤶�</U></td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(59)" size="6" value="<%=Dvbbs.forum_setting(59)%>">&nbsp;byte
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�Զ���ͷ�����ٷ�����������</U><BR>��������������Ϊ0</td>
<td width="50%" class=ForumRowHighlight>  
<input type="text" name="forum_setting(60)" size="6" value="<%=Dvbbs.forum_setting(60)%>">&nbsp;ƪ
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Զ���ͷ��ע����������</U><BR>��������������Ϊ0</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(61)" size="6" value="<%=Dvbbs.forum_setting(61)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>�Զ���ͷ������������������һ������</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(62)" value=0 <%if Dvbbs.forum_setting(62)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(62)" value=1 <%if Dvbbs.forum_setting(62)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Զ���ͷ����Ҫ���εĴ���</U><BR>ÿ�������ַ��á�|�����Ÿ���</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(63)" size="50" value="<%=Dvbbs.forum_setting(63)%>">
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>������ʾҳ���Ƿ���ʾ֧�����͸��û�����ͼ��</U></td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(89)" value=0 <%if Dvbbs.forum_setting(89)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(89)" value=1 <%if Dvbbs.forum_setting(89)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting17"></a><b>��ˢ�»���</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight> <U>��ˢ�»���</U><BR>��ѡ�������д���������ˢ��ʱ��<BR>�԰����͹���Ա��Ч</td>
<td width="50%" class=ForumRowHighlight>  
<input type=radio name="forum_setting(19)" value=0 <%if Dvbbs.forum_setting(19)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(19)" value=1 <%if Dvbbs.forum_setting(19)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>���ˢ��ʱ����</U><BR>��д����Ŀ��ȷ�������˷�ˢ�»���<BR>���������б����ʾ����ҳ��������</td>
<td width="50%" class=Forumrow>  
<input type="text" name="forum_setting(20)" size="3" value="<%=Dvbbs.forum_setting(20)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=ForumRowHighlight><U>��ˢ�¹�����Ч��ҳ��</U><BR>��ȷ�������˷�ˢ�¹���<BR>��ָ����ҳ�潫�з�ˢ�����ã��û����޶���ʱ���ڲ����ظ��򿪸�ҳ�棬����һ��������Դ���ĵ�����<BR>ÿ��ҳ�������á�|�����Ÿ���</td>
<td width="50%" class=ForumRowHighlight> 
<input type="text" name="forum_setting(64)" size="50" value="<%=Dvbbs.forum_setting(64)%>">
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting20"></a><b>����ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td class=Forumrow width="50%"><U>ÿ������ʱ����</U></td>
<td class=Forumrow width="43%"> 
<input type="text" name="Forum_Setting(3)" size="6" value="<%=Dvbbs.Forum_Setting(3)%>">&nbsp;��
</td>
<input type="hidden" id="s_1" value="<b>ÿ������ʱ����</b><br><li>���ú����ÿ������ʱ���������Ա����û�����������ͬ���������Ĵ�����̳��Դ">
<td class=forumRow><a href=# onclick="helpscript(s_1);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight><U>�����ִ���С����󳤶�</U><BR>��С������ַ����÷��š�|���ָ�����λΪ�ֽ�<BR>��С�ַ��������ù�С������ַ��������ù��󣬽�����Ĭ��ֵ</td>
<td class=ForumrowHighLight > 
<input type="text" name="Forum_Setting(4)" size="8" value="<%=Dvbbs.Forum_Setting(4)%>">
</td>
<input type="hidden" id="s_2" value="<b>�����ִ���С����󳤶�</b><br><li>��С������ַ����÷��š�|���ָ�����λΪ�ֽ�<br><li>��С�ַ��������ù�С������ַ��������ù������ù�С���߹��󶼽����Ĵ�����̳��Դ">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_2);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow ><U>�������Բ����ִ��������ƵĴ�</U><BR>ÿ���ַ����÷��š�|���ָ�</td>
<td class=Forumrow> 
<input type="text" name="Forum_Setting(9)" size="50" value="<%=Dvbbs.Forum_Setting(9)%>">&nbsp;
</td>
<input type="hidden" id="s_3" value="<b>�������Բ����ִ��������ƵĴ�</b><br><li>ÿ���ַ����÷��š�|���ָ�<br><li>�������д�����ִ��������ƵĴʣ�����ʹһЩ�����Ҽ򵥵ĵ������������������ͬʱ���뿼�������ִ����ȵĳ����Ǻ����ĵ���Դ�����ȵ�">
<td class=Forumrow><a href=# onclick="helpscript(s_3);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight><U>�����������Ľ����</U><BR>���鲻Ҫ���ù���</td>
<td class=ForumrowHighLight> 
<input type="text" name="Forum_Setting(12)" size="6" value="<%=Dvbbs.Forum_Setting(12)%>">&nbsp;��
</td>
<input type="hidden" id="s_4" value="<b>�����������Ľ����</b><br><li>��λΪ����<br><li>���������Ľ���������ĵ���Դ�����ȣ����������">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_4);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow>
<U>�����������������ж�Ӧ���������������������׼</U><BR>��������������������÷��š�|���ָ�����λΪ����<BR>���������������ù��󣬽�����Ĭ��ֵ</td>
<td class=Forumrow> 
<input type="text" name="Forum_Setting(13)" size="8" value="<%=Dvbbs.Forum_Setting(13)%>">
</td>
<input type="hidden" id="s_5" value="<b>�����������������ж�Ӧ���������������������׼</b><br><li>��������������������÷��š�|���ָ�����λΪ����<br><li>��Ϊ����������������������������׼����̳��Դ���ĳ����ȣ����������">
<td class=Forumrow><a href=# onclick="helpscript(s_5);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight> <U>�Ƿ���ȫ������</U><BR>ACCESS���ݿⲻ���鿪��<BR>SQL���ݿ�����ȫ���������Կ���</td>
<td class=ForumrowHighLight>  
<input type=radio name="Forum_Setting(16)" value=0 <%If Dvbbs.Forum_Setting(16)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(16)" value=1 <%If Dvbbs.Forum_Setting(16)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="s_6" value="<b>�Ƿ���ȫ������</b><br><li>ACCESS���ݿ������������ϴ�����¿������������Ĵ�����Դ��SQL���ݿ⿪�����ݿ�ȫ���������ʹ�ñ�ѡ��<br><li>����SQL���ݿ��ȫ�������뿴΢����ذ����ĵ�">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_6);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow> <U>�û��б������û�������</U></td>
<td class=Forumrow>  
<input type=radio name="Forum_Setting(17)" value=0 <%if Dvbbs.Forum_Setting(17)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(17)" value=1 <%if Dvbbs.Forum_Setting(17)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="s_7" value="<b>�û��б������û�������</b><br><li>��������Ŀ�����û��б��п��Զ��û�����������<br><li>�����û����ݰ�ȫ�ϵĿ��ǣ���Ҳ���Թرո�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(s_7);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight> <U>�û��б������г������Ŷ�</U></td>
<td class=ForumrowHighLight>  
<input type=radio name="Forum_Setting(18)" value=0 <%if Dvbbs.Forum_Setting(18)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(18)" value=1 <%if Dvbbs.Forum_Setting(18)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="s_8" value="<b>�û��б������г������Ŷ�</b><br><li>��������Ŀ�����û��б��п����г���̳�еĹ����Ŷ����ϣ��������������ϵȼ����û�<br><li>�����û����ݰ�ȫ�ϵĿ��ǣ���Ҳ���Թرո�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_8);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow> <U>�û��б������г������û�</U></td>
<td class=Forumrow>  
<input type=radio name="Forum_Setting(27)" value=0 <%if Dvbbs.Forum_Setting(27)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(27)" value=1 <%if Dvbbs.Forum_Setting(27)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="s_9" value="<b>�û��б������г������û�</b><br><li>��������Ŀ�����û��б��п����г���̳�е����е��û�����<br><li>�����û����ݰ�ȫ�ϵĿ��ǣ���Ҳ���Թرո�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(s_9);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=ForumrowHighLight> <U>�û��б������г�TOP�����û�</U></td>
<td class=ForumrowHighLight>  
<input type=radio name="Forum_Setting(31)" value=0 <%if Dvbbs.Forum_Setting(31)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(31)" value=1 <%if Dvbbs.Forum_Setting(31)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="s_10" value="<b>�û��б������г�TOP�����û�</b><br><li>��������Ŀ�����û��б��п����г���̳���շ����ͻ��������û�����<br><li>�����û����ݰ�ȫ�ϵĿ��ǣ���Ҳ���Թرո�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(s_10);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td class=Forumrow><U>�û��б�TOP����</U></td>
<td class=Forumrow> 
<input type="text" name="forum_setting(68)" size="6" value="<%=Dvbbs.forum_setting(68)%>">&nbsp;��
</td>
<input type="hidden" id="s_11" value="<b>�û��б�TOP����</b><br><li>�ڿ�����TOP���е�����£����������������õ����ֶ�ȡ���涨��Ŀ���û�����<br><li>�����û����ݰ�ȫ�ϵĿ��Ǻͳ�����̳��Դ���ķ���Ŀ��ǣ���Ҳ���Լ��ٸ�ѡ���������Ŀ">
<td class=Forumrow><a href=# onclick="helpscript(s_11);return false;" class="helplink"><img src="images/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting18"></a><b>��̳��ҳ����</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>ÿҳ��ʾ����¼</U><BR>������̳���кͷ�ҳ�йص���Ŀ�������б��������ӳ��⣩</td>
<td class=Forumrow  width="50%">  
<input type="text" name="forum_setting(11)" size="3" value="<%=Dvbbs.forum_setting(11)%>">&nbsp;��
</td>
</tr>
</table><BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting16"></a><b>����ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr>
<td class=Forumrow  width="50%"> <U>��Ϊ���Ż�����������ֵ</U><BR>��׼Ϊ����ظ���</td>
<td class=Forumrow  width="50%">  
<input type="text" name="forum_setting(44)" size="3" value="<%=Dvbbs.forum_setting(44)%>">&nbsp;��
</td>
</tr>
<tr> 
<td class=ForumRowHighlight> <U>�༭����������ʾ"��xxx��yyy�༭"����Ϣ</U></td>
<td class=ForumRowHighlight>  
<input type=radio name="forum_setting(48)" value=0 <%if Dvbbs.forum_setting(48)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(48)" value=1 <%if Dvbbs.forum_setting(48)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow> <U>����Ա�༭����ʾ"��XXX�༭"����Ϣ</U></td>
<td class=Forumrow>  
<input type=radio name="forum_setting(49)" value=0 <%if Dvbbs.forum_setting(49)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(49)" value=1 <%if Dvbbs.forum_setting(49)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=ForumRowHighlight> <U>�ȴ�"��XXX�༭"��Ϣ��ʾ��ʱ��</U><BR>�����û��༭�Լ������Ӷ��������ӵײ���ʾ"��XXX�༭"��Ϣ��ʱ��(�Է���Ϊ��λ)</td>
<td class=ForumRowHighlight>  
<input type="text" name="forum_setting(50)" size="3" value="<%=Dvbbs.forum_setting(50)%>">&nbsp;����
</td>
</tr>
<tr> 
<td class=Forumrow> <U>�༭����ʱ��</U><BR>�༭�������ӵ�ʱ������(�Է���Ϊ��λ, 1����1440����) �������ʱ������, ֻ�й���Ա�Ͱ������ܱ༭��ɾ������. �������ʹ�������, ������Ϊ0</td>
<td class=Forumrow>  
<input type="text" name="forum_setting(51)" size="3" value="<%=Dvbbs.forum_setting(51)%>">&nbsp;����
</td>
</tr>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>��������</b>[<a href="#top">����</a>]
</tr>
<tr> 
<td class=ForumRowHighlight  width="50%"> <U>�Ƿ�����̳����</U></td>
<td class=ForumRowHighlight>  
<input type=radio name="forum_setting(32)" value=0 <%if Dvbbs.forum_setting(32)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(32)" value=1 <%if Dvbbs.forum_setting(32)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>VIP�û��鹦�ܿ�������</b>[<a href="#top">����</a>]
</tr>
<tr> 
<td class=ForumRowHighlight  width="50%"><a name="SettingVIP"></a>
<U>�Ƿ�VIP�û��鹦��</U>
<br>������VIP�û��鹦�ܣ���ȷ����̳�û���(�ȼ�)�������Ƿ������VIP�û��顣
</td>
<td class=ForumRowHighlight>  
<input type=radio name="forum_setting(43)" value=0 <%if Dvbbs.forum_setting(43)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(43)" value=1 <%if Dvbbs.forum_setting(43)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</table>
<BR>
<table border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="settingxu"></a><b>�����ٷ����ѡ��</b>[<a href="#top">����</a>]
</tr>
<tr> 
<td class=Forumrow width="50%"> <U>���߹����ܿ���</U></td>
<td class=Forumrow width="43%">  
<input type=radio name="forum_setting(90)" value=0 <%if Dvbbs.forum_setting(90)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(90)" value=1 <%if Dvbbs.forum_setting(90)="1" then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td class=ForumRowHighlight> <U>����������������</U></td>
<td class=ForumRowHighlight>  
<input type=radio name="forum_setting(91)" value=0 <%if Dvbbs.forum_setting(91)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(91)" value=1 <%if Dvbbs.forum_setting(91)="1" then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow> <U>�������Ĳ��ö������ݿ�</U><BR>(����Ϊ�������������޸�CONN.ASP�ļ������ö������ݿ�·��)</td>
<td class=Forumrow>  
<input type=radio name="forum_setting(92)" value=0 <%if Dvbbs.forum_setting(92)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(92)" value=1 <%if Dvbbs.forum_setting(92)="1" then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td class=ForumRowHighlight width="50%"> <U>ħ�����飨ͷ���ܿ���</U><BR>�ù������ݿ���õ����������ݿ⣬���ܿɶ����ڵ�������֮�⿪��</td>
<td class=ForumRowHighlight width="43%">  
<input type=radio name="forum_setting(98)" value=0 <%if Dvbbs.forum_setting(98)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(98)" value=1 <%if Dvbbs.forum_setting(98)="1" then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>�Ƿ����ò��͹���</U><BR>���������boke/config.asp�ļ������������</td>
<td class=Forumrow>  
<input type=radio name="forum_setting(99)" value=0 <%if Dvbbs.forum_setting(99)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(99)" value=1 <%if Dvbbs.forum_setting(99)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=Forumrow  width="50%"> <U>�Ƿ�������������</U>(��δ��ͨ)</td>
<td class=Forumrow>  
<input type=radio name="forum_setting(82)" value=1 <%if Dvbbs.forum_setting(82)="1" then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(82)" value=0 <%if Dvbbs.forum_setting(82)="0" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting23"></a><b>��̳��һ�������</b></th>
</tr>
<tr> 
<td class=ForumRowHighlight width="50%"> <U>��Ǯ���һ���</U></td>
<td class=ForumRowHighlight>  
<input type="text" name="forum_setting(93)" size="6" value="<%=Dvbbs.forum_setting(93)%>">&nbsp;��Ǯ=1���
</td>
</tr>
<tr> 
<td class=Forumrow width="50%"> <U>�������һ���</U></td>
<td class=Forumrow>  
<input type="text" name="forum_setting(94)" size="6" value="<%=Dvbbs.forum_setting(94)%>">&nbsp;����=1���
</td>
</tr>
<tr> 
<td class=ForumRowHighlight width="50%"> <U>�������һ���</U></td>
<td class=ForumRowHighlight>  
<input type="text" name="forum_setting(95)" size="6" value="<%=Dvbbs.forum_setting(95)%>">&nbsp;����=1���
</td>
</tr>
<tr>
<td class=Forumrow width="50%"> <U>��ȯ���һ���</U></td>
<td class=Forumrow>  
<input type="text" name="forum_setting(96)" size="6" value="<%=Dvbbs.forum_setting(96)%>">&nbsp;��ȯ=1���
</td>
</tr>
<tr> 
<th height=25 colspan=3 align=left id=tabletitlelink><a name="setting23"></a><b>��������</b></th>
</tr>
<tr>
<td class=ForumRowHighlight width="50%"> <U>����ÿ�տɽ�����Ҹ���</U></td>
<td class=ForumRowHighlight>  
<input type="text" name="forum_setting(97)" size="6" value="<%=Dvbbs.forum_setting(97)%>">&nbsp;�����
</td>
</tr>
<tr> 
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<SCRIPT LANGUAGE="JavaScript">
CheckSel('forum_setting(2)','<%=Dvbbs.forum_setting(2)%>');
CheckSel('UploadSetting(0)','<%=UploadSetting(0)%>');
CheckSel('UploadSetting(2)','<%=UploadSetting(2)%>');
CheckSel('UploadSetting(3)','<%=UploadSetting(3)%>');
CheckSel('UploadSetting(7)','<%=UploadSetting(7)%>');
CheckSel('UploadSetting(13)','<%=UploadSetting(13)%>');
CheckSel('UploadSetting(16)','<%=UploadSetting(16)%>');
CheckSel('UploadSetting(17)','<%=UploadSetting(17)%>');
</script>
<div id="Issubport0" style="display:none">��ѡ��EMAIL�����</div>
<div id="Issubport999" style="display:none"></div>
<%
Dim InstalledObjects(12)
InstalledObjects(1) = "JMail.Message"				'JMail 4.3
InstalledObjects(2) = "CDONTS.NewMail"				'CDONTS
InstalledObjects(3) = "Persits.MailSender"			'ASPEMAIL
'-----------------------
InstalledObjects(4) = "Adodb.Stream"				'Adodb.Stream
InstalledObjects(5) = "Persits.Upload"				'Aspupload3.0
InstalledObjects(6) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
InstalledObjects(7) = "DvFile.Upload"				'DvFile-Up V1.0
'-----------------------
InstalledObjects(9) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
InstalledObjects(10) = "Persits.Jpeg"				'AspJpeg
InstalledObjects(11) = "SoftArtisans.ImageGen"		'SoftArtisans ImgWriter V1.21
InstalledObjects(12) = "sjCatSoft.Thumbnail"		'sjCatSoft.Thumbnail V2.6

For i=1 to 12
	Response.Write "<div id=""Issubport"&i&""" style=""display:none"">"
	If IsObjInstalled(InstalledObjects(i)) Then Response.Write InstalledObjects(i)&":<font color=red><b>��</b>������֧��!</font>" Else Response.Write InstalledObjects(i)&"<b>��</b>��������֧��!" 
	Response.Write "</div>"
Next
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function chkselect(s,divid)
{
var divname='Issubport';
var chkreport;
	s=Number(s)
	if (divid=="know1")
	{
	divname=divname+s;
	}
	if (divid=="know2")
	{
	s+=4;
	if (s==1003){s=999;}
	divname=divname+s;
	}
	if (divid=="know3")
	{
	s+=9;
	if (s==1008){s=999;}
	divname=divname+s;
	}
document.getElementById(divid).innerHTML=divname;
chkreport=document.getElementById(divname).innerHTML;
document.getElementById(divid).innerHTML=chkreport;
}
//-->
</SCRIPT>
<%
end sub

Sub saveconst()
	Dim Forum_copyright,Forum_info,forum_setting,iforum_setting,isetting
	Dim Forum_Maxonline,Forum_TopicNum,Forum_PostNum
	Dim Forum_UserNum,Forum_MaxPostNum,Forum_MaxPostDate,Forum_MaxonlineDate
	Dim Forum_pack
	Dim UploadSetting,Tempstr
	
	If not IsDate(Request.Form("Forum_Setting(74)")) Then 
		Errmsg=ErrMsg + "<li>��̳�������ڱ�����һ����Ч���ڡ�"
		Dvbbs_error()
		Exit Sub
	End If
	
	If not IsDate(Request.Form("Forum_MaxPostDate")) Then 
		Errmsg=ErrMsg + "<li>��̳����շ�������ʱ�����ڱ�����һ����Ч���ڡ�"
		Dvbbs_error()
		Exit Sub
	Else
		Forum_MaxPostDate=Request.Form("Forum_MaxPostDate")
	End If
	
	If not IsDate(Request.Form("Forum_MaxonlineDate")) Then 
		Errmsg=ErrMsg + "<li>��ʷ���ͬʱ���߼�¼����ʱ�����ڱ�����һ����Ч���ڡ�"
		Dvbbs_error()
		Exit Sub
	Else
		Forum_MaxonlineDate=Request.Form("Forum_MaxonlineDate")
	End If
	
	Forum_Maxonline	= Request.Form("Forum_Maxonline")
	Forum_TopicNum	= Request.Form("Forum_TopicNum")
	Forum_PostNum	= Request.Form("Forum_PostNum")
	Forum_UserNum	= Request.Form("Forum_UserNum")
	Forum_MaxPostNum= Request.Form("Forum_MaxPostNum")
	Forum_pack	= Request.Form("Forum_pack(0)")&"|||"&Trim(Request.Form("Forum_pack(1)"))
	
	If Not ISNumeric(Forum_Maxonline&Forum_TopicNum&Forum_PostNum&Forum_UserNum&Forum_MaxPostNum) Then 
		Errmsg=ErrMsg + "<li>�Ƿ��Ĳ�������̳ϵͳ���ݳ����ύ��ֹ��"
		Dvbbs_error()
		Exit Sub
	End If
	
	If not isnumeric(request.Form("cid")) or not isnumeric(request.Form("Sid")) Then
		Errmsg=ErrMsg + "<li>��ѡ��ģ������"
		Dvbbs_error()
		Exit Sub
	End IF
	
	UploadSetting = ""
	For i=0 To 20
		Tempstr = Trim(Request.Form("UploadSetting("&i&")"))
		If Tempstr = "" Then
			UploadSetting = UploadSetting & 0
		Else
			UploadSetting = UploadSetting & Replace(Replace(Tempstr,"|",""),",","")
		End If
		If i<20 Then
			UploadSetting = UploadSetting & "|"
		End If
	Next
	
	Dim setingdata,j
	If Forum_Maxonline="" Then Forum_Maxonline=0
	If Forum_TopicNum="" Then Forum_TopicNum=0
	If Forum_PostNum="" Then Forum_PostNum=0
	If Forum_UserNum="" Then Forum_UserNum=0
	If Forum_MaxPostNum="" Then Forum_MaxPostNum=0
	For i = 0 To 120
		If Trim(request.Form("Forum_Setting("&i&")"))="" Or i=70 or i = 7 Then
			'Response.Write "Forum_Setting("&i&")<br>"
			isetting=0
			If i=7 Then
				isetting = UploadSetting
			End If
			If i=70 Then
				isetting=""
				For j=0 to  23
					If isetting="" Then
						If Request.form("Forum_Setting(70)"&j)="1" Then
							isetting="1"
						Else
							isetting="0"
						End If
					Else
						If Request.form("Forum_Setting(70)"&j)="1" Then
							isetting=isetting&"|1"
						Else
							isetting=isetting&"|0"
						End If
					End If
				Next
			End If		
		Else
			isetting=Replace(Trim(request.Form("Forum_Setting("&i&")")),",","")
		End If
	
		If i = 0 Then
			forum_setting = isetting
		Else
			forum_setting = forum_setting & "," & isetting
		End If
	Next
	
	For i = 0 To 13
		If Trim(Request.Form("Forum_info("&i&")")) = "" And i <> 4 And i <> 12 And i<>13 Then
			'Response.Write "Forum_info("&i&")<br>"
			isetting=0
		Else
			isetting=Replace(Trim(request.Form("Forum_info("&i&")")),",","")
		End If
		If i = 0 Then
			Forum_info = isetting
		Else
			Forum_info = Forum_info & "," & isetting
		End If
	Next
	Forum_copyright=request("copyright")
	'forum_info|||forum_setting|||forum_user|||copyright|||splitword|||stopreadme
	Set rs=Dvbbs.execute("select forum_setting from dv_setup")
	iforum_setting=split(rs(0),"|||")
	forum_setting=forum_info & "|||" & forum_setting & "|||" & iforum_setting(2) & "|||" & Forum_copyright & "|||" & iforum_setting(4) & "|||" & request.Form("StopReadme")
	forum_setting=Replace(forum_setting,"'","''")
	Dim Forum_AdminFolder
	Forum_AdminFolder=Request("Forum_AdminFolder")
	If Forum_AdminFolder<>"" Then
		Forum_AdminFolder=Forum_AdminFolder&"/"
		Forum_AdminFolder=replace(Forum_AdminFolder,"\","/")
		Forum_AdminFolder=replace(Forum_AdminFolder,"//","/")
	Else
		Forum_AdminFolder=Dvbbs.CacheData(33,0)
	End If
	'Response.Write forum_setting
	'Response.end
	
	sql="update Dv_setup set Forum_AdminFolder='"&Dvbbs.Checkstr(Forum_AdminFolder)&"',Forum_Setting='"&forum_setting&"',forum_sid="&Dvbbs.checkstr(request.Form("Sid"))&",forum_Cid="&Dvbbs.checkstr(request.Form("cid"))
	sql=sql+",Forum_Maxonline="&Forum_Maxonline&",Forum_TopicNum="&Forum_TopicNum&",Forum_PostNum="&Forum_PostNum &",Forum_UserNum="&Forum_UserNum&",Forum_MaxPostNum="&Forum_MaxPostNum&",Forum_MaxPostDate='"&Forum_MaxPostDate &"',Forum_MaxonlineDate='"&Forum_MaxonlineDate&"',Forum_pack='"&Forum_pack&"'"
	dvbbs.execute(sql)
	Dvbbs.loadSetup()
	Dv_suc("������̳������Ϣ�ɹ�")
end sub

'�ָ�Ĭ������
Sub restore()
	Dim Forum_setting
	forum_setting="  �����ȷ���̳,http://bbs.dvbbs.net,�����ȷ�,http://www.aspsky.net/,,eway@aspsky.net,images/logo.gif,http://www.aspsky.cn/email.asp,aspsky|dvbbs|����|������̳|asp|��̳|���,����ʱ��,������̳��ʹ������ࡢ�������������������̳��Ҳ�ǹ���֪���ļ�������վ�㣬ϣ�����������Ŭ������Ϊ�������ܶ෽��,index.asp,,|||0,300,0,60,2|20,1,1,2|100|0|999|bbs.dvbbs.net|12|#FF0000|Arial|0|images/WaterMap.gif|0.7|110|35|4|120|100|1|2|#0066cc|0|0,40,dvbbs|sql|aspsky|asp|php|cgi|jsp|htm,0,20,500,20|200,0,0,1,1,1,0,30,0,7200,0,1,0,0,1,1,0,1,1,1,1,1,1,0,1,32,32,1,10,1,0,10,0,1,0,1,1,0,0,0,1,10,1,0,120,30,9,15,4,0,0,0,1,0,1,20,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,0,0,0,2000-3-26,0,UploadFile/,object|EMBED|,1,1,1,1,0,0,5,0,0,0,0,0,1,1,0,50,30,20,5,100,1,0,0|||100,15,12,12,8,60,10,5,5,6,30,6,2,2,3,25,15,20|||Copyright &copy;2000 - 2005  <a href=http://www.dvbbs.net><font face=Verdana, Arial, Helvetica, sans-serif><b>Aspsky<font color=#CC0000>.Net</font></b></font></a>|||!,@,#,$,%,^,&,*,(,),{,},[,],|,\,.,/,?,`,~|||��̳��ͣʹ��"
	Conn.Execute("update Dv_setup set Forum_Setting='"&forum_setting&"'")
	Dv_suc("��ԭ��̳�������óɹ�")
	Dvbbs.loadSetup()
End Sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>