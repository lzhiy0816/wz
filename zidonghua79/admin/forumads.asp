<!--#include file =../conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="../inc/dv_clsother.asp" -->
<%
Call Head()
Dim admin_flag
admin_flag=",2,"
If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=../admin_login.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	dvbbs_error()
Else 
	If request("action")="save" Then 
		Call saveconst()
	Else
		Call consted()
	End If
		If founderr then call dvbbs_error()
		footer()
	End If 

Sub consted()
dim sel
%>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">

<tr> 
<th height="23" colspan="2" class="tableHeaderText"><b>��̳�������</b>����Ϊ���÷���̳�����Ƿ���̳��ҳ��棬����ҳ��Ϊ������ʾҳ�棩</th>
</tr>
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2><B>˵��</B>��<BR>1����ѡ����ѡ���Ϊ��ǰ��ʹ������ģ�壬����ɲ鿴��ģ�����ã�������ģ��ֱ�Ӳ鿴��ģ�岢�޸����á������Խ�����������ñ����ڶ����̳������<BR>2����Ҳ���Խ������趨����Ϣ���沢Ӧ�õ�����ķ���̳���������У��ɶ�ѡ<BR>3�����������һ���������ñ�İ�������ã�ֻҪ����ð������ƣ������ʱ��ѡ��Ҫ���浽�İ����������Ƽ��ɡ�
<hr size=1 width="100%" color=blue>
</td>
</tr>
<FORM METHOD=POST ACTION="">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
�鿴�ְ��������ã���ѡ�������������Ӧ����&nbsp;&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option value="">�鿴�ְ�������ѡ��</option>
<%
Dim ii
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value=""forumads.asp?boardid="&rs(0)&""">"
Select Case rs(2)
	Case 0
		Response.Write "��"
	Case 1
		Response.Write "&nbsp;&nbsp;��"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;��"
	Next
	Response.Write "&nbsp;&nbsp;��"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
</tr>
</FORM>
</table><BR>
<form method="POST" action=forumads.asp?action=save>
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td width="100%" class="forumRowHighlight" colspan=2>
<input type=checkbox name="getskinid" value="1" <%if request("getskinid")="1" or request("boardid")="" then Response.Write "checked"%>><a href="forumads.asp?getskinid=1">��̳Ĭ�Ϲ��</a><BR> ����˴�������̳Ĭ�Ϲ�����ã�Ĭ�Ϲ�����ð�������<FONT COLOR="blue">��</FONT>��������������ݣ��������б�������ʾ�����澫�������淢���ȣ�<FONT COLOR="blue">����</FONT>��ҳ�档<hr size=1 width="90%" color=blue>
</td>
</tr>
<tr> 
<td width="200" class="forumrow" valign=top>
�����汣��ѡ��<BR>
�밴 CTRL ����ѡ<BR>
<select name="getboard" size="28" style="width:100%" multiple>
<%
set rs=Dvbbs.Execute("select boardid,boardtype,depth from dv_board order by rootid,orders")
do while not rs.eof
Response.Write "<option "
if rs(0)=dvbbs.boardid then
Response.Write " selected"
end if
Response.Write " value="&rs(0)&">"
Select Case rs(2)
	Case 0
		Response.Write "��"
	Case 1
		Response.Write "&nbsp;&nbsp;��"
End Select
If rs(2)>1 Then
	For ii=2 To rs(2)
		Response.Write "&nbsp;&nbsp;��"
	Next
	Response.Write "&nbsp;&nbsp;��"
End If
Response.Write rs(1)
Response.Write "</option>"
rs.movenext
loop
rs.close
set rs=nothing
%>
</select>
</td>
<td class="forumrow" valign=top>
<table>
<tr>
<td width="200" class="forumrow"><B>��ҳ����������</B><BR>��������˻�����湦���еĶ�����棬�˴�����Ϊ��Ч</td>
<td width="*" class="forumrow"> 
<textarea name="Forum_ads(0)" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(0))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��ҳβ��������</B></td>
<td width="*" class="forumrow"> 
<textarea name="Forum_ads(1)" cols="50" rows="3"><%=server.htmlencode(Dvbbs.Forum_ads(1))%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>������ҳ�������</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(2)" value=0 <%if Dvbbs.Forum_ads(2)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_ads(2)" value=1 <%if Dvbbs.Forum_ads(2)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ�������ͼƬ��ַ</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(3)" size="35" value="<%=Dvbbs.Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ����������ӵ�ַ</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(4)" size="35" value="<%=Dvbbs.Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ�������ͼƬ���</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(5)" size="3" value="<%=Dvbbs.Forum_ads(5)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ�������ͼƬ�߶�</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(6)" size="3" value="<%=Dvbbs.Forum_ads(6)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>������ҳ���¹̶����</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(13)" value=0 <%if Dvbbs.Forum_ads(13)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_ads(13)" value=1 <%if Dvbbs.Forum_ads(13)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ��ַ</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(8)" size="35" value="<%=Dvbbs.Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶�������ӵ�ַ</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(9)" size="35" value="<%=Dvbbs.Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ���</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(10)" size="3" value="<%=Dvbbs.Forum_ads(10)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>��̳��ҳ���¹̶����ͼƬ�߶�</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(11)" size="3" value="<%=Dvbbs.Forum_ads(11)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>�Ƿ�������������</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(7)" value=0 <%if Dvbbs.Forum_ads(7)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_ads(7)" value=1 <%if Dvbbs.Forum_ads(7)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="*" class="forumrow" valign="top" colspan=2><B>��̳�������������</B> <br>֧��HTML�﷨��ÿ��������һ�У��ûس��ֿ���</td>
</tr>
<tr>
<%
Dim Ads_14
If UBound(Dvbbs.Forum_ads)>13 Then
	Ads_14=Dvbbs.Forum_ads(14)
End If
%>
<td width="*" class="forumrow" colspan=2> 
<textarea name="Forum_ads(14)" style="width:100%" rows="10"><%=Ads_14%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow"><B>�Ƿ���ҳ�����ֹ��λ</B></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(12)" value=0 <%if Dvbbs.Forum_ads(12)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_ads(12)" value=1 <%if Dvbbs.Forum_ads(12)="1" then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr>
<%
Dim Ads_15
If UBound(Dvbbs.Forum_ads)>14 Then
	Ads_15=Dvbbs.Forum_ads(15)
End If
%>
<td width="200" class="forumrow"><B>ҳ�����ֹ��λ����(����)</B><BR>��ȷ���Ѵ���ҳ�����ֹ��λ����<BR></td>
<td width="*" class="forumrow"> 
<input type=radio name="Forum_ads(15)" value=0 <%if Ads_15="0" then%>checked<%end if%>>�����б�&nbsp;
<input type=radio name="Forum_ads(15)" value=1 <%if Ads_15="1" then%>checked<%end if%>>��������&nbsp;
<input type=radio name="Forum_ads(15)" value=2 <%if Ads_15="2" then%>checked<%end if%>>���߶���ʾ&nbsp;
<input type=radio name="Forum_ads(15)" value=3 <%if Ads_15="3" then%>checked<%end if%>>���߶�����ʾ&nbsp;
</td>
</tr>
<tr>
<%
Dim Ads_17
If UBound(Dvbbs.Forum_ads)>16 Then
	Ads_17=Dvbbs.Forum_ads(17)
End If
%>
<td width="200" class="forumrow"><B>���ֹ��ÿ�й�����</B></td>
<td width="*" class="forumrow"> 
<input type="text" name="Forum_ads(17)" size="3" value="<%=Ads_17%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="*" class="forumrow" valign="top" colspan=2><B>ҳ�����ֹ��λ����</B> <br>֧��HTML�﷨��ÿ�����һ�У��ûس��ֿ���</td>
</tr>
<tr>
<td width="*" class="forumrow" colspan=2>
<%
Dim Ads_16
If UBound(Dvbbs.Forum_ads)>15 Then
	Ads_16=Dvbbs.Forum_ads(16)
End If
%>
<textarea name="Forum_ads(16)" style="width:100%" rows="10"><%=Ads_16%></textarea>
</td>
</tr>
<tr> 
<td width="200" class="forumrow">&nbsp;</td>
<td width="*" class="forumrow"> 
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
<%
end sub

Sub SaveConst()
	Dim iSetting
	For i = 0 To 30
		If Trim(Request.Form("Forum_ads("&i&")"))="" Then
			If i = 1 Or i = 0 Then
				iSetting = ""
			ElseIf i = 17 Then
				iSetting = 1
			Else
				iSetting = 0
			End If
		Else
			iSetting=Replace(Trim(Request.Form("Forum_ads("&i&")")),"$","")
		End If

		If i = 0 Then
			Dvbbs.Forum_ads = iSetting
		Else
			Dvbbs.Forum_ads = Dvbbs.Forum_ads & "$" & iSetting
		End If
	Next

	If Request("getskinid")="1" Then
		Sql = "Update Dv_Setup Set Forum_ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"'"
		Dvbbs.Execute(sql)
		Dvbbs.loadSetup()
	End If
	If Request("getboard")<>"" Then
		Sql = "Update Dv_Board Set Board_Ads='"&Replace(Dvbbs.Forum_ads,"'","''")&"' Where BoardID In ("&Request("getboard")&")"
		Dvbbs.Execute(Sql)
		Dvbbs.loadSetup()
		RestoreBoardCache()
		Dvbbs.ReloadBoardCache request("getboard")
	Else
		Dvbbs.loadSetup()
		RestoreBoardCache()
	End If

	Dv_suc("������óɹ���")
End Sub
Sub RestoreBoardCache()
	Dim Board,node
	Dvbbs. LoadBoardList()
	For Each node in Application(Dvbbs.CacheName &"_style").documentElement.selectNodes("style/@id")
		Application.Contents.Remove(Dvbbs.CacheName & "_showtextads_"&node.text)
		For Each board in Application(Dvbbs.CacheName&"_boardlist").documentElement.selectNodes("board/@boardid")
			Dvbbs.LoadBoardData board.text
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text)
			Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_"& board.text &"_"&node.text&"_-time")
		Next
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text)
		Application.Contents.Remove(dvbbs.CacheName & "_Text_ad_0_"& node.text&"_-time")
	Next
End Sub
%>