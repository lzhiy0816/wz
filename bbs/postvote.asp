<!-- #include file="setup.asp" -->
<%

id=int(Request("id"))

if Request("menu")="look" then

list=""&Application("vote"&id&"")&""
if list=empty then
error2("目前暂时没有记录投票的用户名单！")
else
error2("参与投票的用户名单：\n"&list&"")
end if
end if


if Request.Cookies("username")=empty then
error2("您必须登录后才能投票")
end if




if Request("postvote")="" then
error2("请选择，你要投票的项目！")
end if


if instr(Application("vote"&id&""),""&Request.Cookies("username")&" ")>0 or instr(Request.Cookies("vote"),""&id&"|")>0 then
error2("你已经投过票了，无需重复投票！")
end if



sql="select * from forum where id="&id&""
rs.Open sql,Conn,1,3
for each ho in request.form("postvote")
pollresult=split(rs("pollresult"),"|")
for i = 0 to ubound(pollresult)
if not pollresult(i)="" then
if cint(ho)=i then
pollresult(i)=pollresult(i)+1
end if
allpollresult=""&allpollresult&""&pollresult(i)&"|"
end if
next

rs("pollresult")=allpollresult
rs("lastname")=Request.Cookies("username")
rs("lasttime")=now()
rs.update
allpollresult=""
next


rs.close

Application("vote"&id&"")=""&Request.Cookies("username")&" "&Application("vote"&id&"")&""
Response.Cookies("vote")=""&Request.Cookies("vote")&""&id&"|"

error2("投票成功!")


%>