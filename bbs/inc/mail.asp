<%
if instr(mailaddress,"@")=0 then
error2("E-mail��ַ��д����")
end if
on error resume next
if selectmail="JMail" then

Set JMail=Server.CreateObject("JMail.Message")
If -2147221005 = Err Then response.write "����������֧�� JMail.Message ������뵽��̨�رշ����ʼ����ܣ�"
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
If -2147221005 = Err Then response.write "����������֧�� CDONTS.NewMail ������뵽��̨�رշ����ʼ����ܣ�"
MailObject.Send smtpmail,mailaddress,mailtopic,body
Set MailObject=nothing
end if
%>