<%
select case Request("menu")

case "skins"
Response.Cookies("skins")=""&Request("no")&""
Response.Cookies("skins").Expires=date+3650
response.redirect "main.asp"

case "eremite"
Response.Cookies("eremite")="1"
Response.Cookies("eremite").Expires=date+3650
response.write "<script>alert('����ǰ��״̬Ϊ������');history.back();</script>"

case "online"
Response.Cookies("eremite")="0"
Response.Cookies("eremite").Expires=date+3650
response.write "<script>alert('����ǰ��״̬Ϊ������');history.back();</script>"

end select
%>