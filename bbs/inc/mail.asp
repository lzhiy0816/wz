<%
if instr(mailaddress,"@")=0 then
error2("E-mail地址填写错误！")
end if
on error resume next
if selectmail="JMail" then

Set JMail=Server.CreateObject("JMail.Message")
If -2147221005 = Err Then response.write "本服务器不支持 JMail.Message 组件！请到后台关闭发送邮件功能！"
JMail.Charset="gb2312"
JMail.AddRecipient mailaddress
JMail.Subject = mailtopic
JMail.Body = body
JMail.From = smtpmail
JMail.MailServerUserName = MailServerUserName
JMail.MailServerPassword = MailServerPassword
JMail.Send smtp
Set JMail=nothing

elseif selectmail="CDONTS" then
Set MailObject = Server.CreateObject("CDONTS.NewMail")
If -2147221005 = Err Then response.write "本服务器不支持 CDONTS.NewMail 组件！请到后台关闭发送邮件功能！"
MailObject.Send smtpmail,mailaddress,mailtopic,body
Set MailObject=nothing
end if
%>