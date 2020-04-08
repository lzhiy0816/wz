<!-- #include file="conn.asp" -->
<%
'=========================================================
' 调用代码
' <script src=new.asp?forumid=1&TopicCount=10&TitleCount=15&timeCount=1></script>
' forumid:    指定论坛的ID
' TopicCount: 显示多少条主题
' TitleCount: 标题长度
' timeCount:  1=显示时间 0=不显示时间
'=========================================================
' Copyright (C) 1998-2004 Yuzi.Net. All rights reserved.
' Web: http://www.yuzi.net,http://www.bbsxp.com
' Email: huangzhiyu@yuzi.net
'=========================================================
forumid=int(Request("forumid"))


if Request("TopicCount")=empty then
TopicCount=10
else
TopicCount=int(Request("TopicCount"))
end if

if Request("TitleCount")=empty then
TitleCount=15
else
TitleCount=int(Request("TitleCount"))
end if


Set conn=Server.CreateObject("ADODB.Connection")
conn.open ConnStr
cluburl=Conn.Execute("Select cluburl From clubconfig")(0)

	if forumid = empty then
		sql="select top "&TopicCount&" id,icon,topic,lasttime from forum where deltopic=0 order by id desc"
	else
		sql="select top "&TopicCount&" id,icon,topic,lasttime from forum where forumid="&forumid&" and deltopic=0 order by id desc"
	end if

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn
	do while Not RS.Eof

	if Request("timeCount")=1 then
	response.write "document.write(' <img src="&cluburl&"/images/brow/"&rs("icon")&".gif> <a href="&cluburl&"/ShowPost.htm?id="&rs("id")&" target=_blank>"&left(rs("Topic"),TitleCount)&"</a> ["&rs("lasttime")&"]<br>');"
	else
	response.write "document.write(' <img src="&cluburl&"/images/brow/"&rs("icon")&".gif> <a href="&cluburl&"/ShowPost.htm?id="&rs("id")&" target=_blank>"&left(rs("Topic"),TitleCount)&"</a><br>');"
	end if

	RS.MoveNext
	Loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
