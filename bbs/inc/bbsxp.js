var i=0;
var ii=3;
var kj=1;
var is_max=1;

function yuzi_img(e, o)
{
var zoom = parseInt(o.style.zoom, 10) || 100;
zoom += event.wheelDelta / 12;
if (zoom > 0) o.style.zoom = zoom + '%';
return false;
}

function min_yuzi(){
if(top!=self){
if(kj==1){parent.main_frame.rows='0,*';}
if(getCookie('frame')=='1'){parent.main_frame.rows='50%,*';}
}
}

function img(){if(top!=self){parent.main_frame.rows='*,20';parent.yuzi_frame.kj=1}}

function mid_f(){
if(top!=self){
if(is_max==0){
parent.main_frame.rows='0,*';is_max =1;
document.mid.title="���»�ԭ";
document.mid.src="images/bbs_mid.gif";
parent.yuzi_frame.kj=0;
}else{
parent.main_frame.rows='50%,*';is_max =0;
document.mid.title="���";
document.mid.src="images/bbs_most.gif";
parent.yuzi_frame.kj=0;
}
}
}

function heartBeat(){
diffY=document.body.scrollTop;
percent=.1*(diffY-lastScrollY);
if(percent>0)percent=Math.ceil(percent);
else percent=Math.floor(percent);
document.all.yuzi.style.pixelTop+=percent;
lastScrollY=lastScrollY+percent;
}


function ShowPost(id,username,content,posttime,honor,userface,sex,birthday,experience,membercode,faction,consort,money,postcount,regtime,userlife,usermail,userhome,sign)
{
document.write("<table class=a2 cellspacing=1 cellpadding=0 width=97% align=center border=0><tr><td valign=top align=left><table style=BORDER-COLLAPSE: cellspacing=0 cellpadding=3 width=100% border=0 height=1%>");
ii++
if(ii==5){ii=3}
document.write("<tr class=a"+ii+"><td valign=top width=160 rowspan=4 height=250>");
document.write("<table cellspacing=0 cellpadding=0 width=90% align=center border=0><tr><td>");
document.write("<table border=0 cellspacing=0 width=95%><tr><td><font style=font-size:10pt><b>"+username+"</b></font><br>"+honor+"</td><td align=right valign=top width=40>");

if (""+sex+""!=''){
document.write("<img src=images/"+sex+".gif>��")
}

document.write(astro(""+birthday+""));
document.write("</td></tr></table>");

if (getCookie('showface')!='0'){
document.write("<table width=90% border=0><tr><td align=middle><img src='"+userface+"' onload='javascript:if(this.width>120)this.width=120;if(this.height>120)this.height=120;'></td></tr></table>")
}

document.write("<br>"+level(experience,membercode,username,moderated)+levelimage+"<br>�ȡ�����:"+levelname+"<br>");
if (""+faction+"" !="") {document.write("�š�����:<a style=cursor:hand onclick=\"javascript:open('faction.asp?menu=look&faction="+faction+"','','toolbar=no')\">"+faction+"</a><br>");}
if (""+consort+"" !="") {document.write("�䡡��ż:"+consort+"<br>");}
document.write("�� �� ֵ:"+experience+"<br>");
document.write("�������:"+money+"<br>");
document.write("�ܷ�����:"+postcount+"<br>");
document.write("ע��ʱ��:"+regtime+"<br>");
document.write("�� �� ֵ:<img border=0 src=images/bar1.gif width="+userlife/2+" height=9 alt="+userlife+"><br>");
if(onlinelist.indexOf("|"+username+"|") == -1 ){
document.write("����״̬:<img border=0 src=images/offline1.gif alt='����'><br>");
}else{
document.write("����״̬:<img border=0 src=images/online1.gif alt='����'><br>");
}

document.write("</td></tr></table></td><td valign=bottom align=middle rowspan=4><table height=100% class=a2 cellspacing=0 cellpadding=0 width=1><tr><td width=1></td></tr></table></td><td valign=top align=left colspan=3><table cellspacing=0 cellpadding=0 width=100% border=0>");
document.write("<tr><td><a onclick=document.location='loading.htm' target=yuzi_frame href='Profile.asp?username="+username+"'><img alt='�鿴"+username+"�ĸ�������' src=images/Profile.gif border=0></a> <a style=cursor:hand onclick=\"javascript:open('friend.asp?menu=post&incept="+username+"','','width=320,height=170')\"><img src=images/pm.gif border=0 alt='���Ͷ�ѶϢ��"+username+"'></a> <a href='friend.asp?menu=add&username="+username+"'><img alt='��"+username+"�������' src=images/friend.gif border=0></a> <a onclick=document.location='loading.htm' target=yuzi_frame href='searchok.asp?search=author&searchxm=username&content="+username+"'><img alt='����"+username+"���������������' src=images/find.gif border=0></a> <a href=mailto:"+usermail+"><img alt='���͵��ʸ�"+username+"' src=images/email.gif border=0></a> ");

if (userhome!="" && userhome!="http://"){
document.write("<a target=_blank href="+userhome+"><img alt=����"+userhome+"����ҳ src=images/homepage.gif border=0></a>");
}
i=i+1
document.write(" <a href=javascript:copyText(document.all.yu"+id+")><img alt=����������� src=images/copy.gif border=0></a> <a onclick=document.location='loading.htm' target=yuzi_frame href='retopic.asp?id="+topicid+"&retopicid="+id+"&quote=1&topic="+topic+"'><img alt=���ûظ�������� src=images/reply.gif border=0></a> <a onclick=document.location='loading.htm' target=yuzi_frame href='retopic.asp?id="+topicid+"&topic="+topic+"'><img src=images/replynow.gif border=0 alt=�ظ��������></a></td><td align=right>No.<font color=red><b>"+i+"</b></font>&nbsp;</td></tr><tr valign=top><td colspan=2 style=word-break:break-all><hr width=100% color=777777 size=1>");

if(badlist.indexOf(username) == -1 ){
document.write(ybbcode("<span id=yu"+id+">"+content+"</span>"));
}else{
document.write("<span id=yu"+id+">==============================<br>������<font color=RED>���û������ѱ����ˡ�����</font><br>==============================</span>")
}
document.write("</td></tr></table></td></tr><tr class=a"+ii+"><td valign=bottom>")
document.write("<table cellspacing=0 cellpadding=0 width=100%><tr><td align=right valign=bottom colspan=3>")
if (getCookie('sign')!='0' && sign!=""){document.write("��������������������<br>"+ybbcode(sign)+"");}
document.write("<hr width=100% color=777777 size=1></td></tr><tr class=a"+ii+"><td valign=bottom rowspan=3><a onclick=document.location='loading.htm' target=yuzi_frame href='EditTopic.asp?id="+topicid+"&retopicid="+id+"&topic="+topic+"'><img src=images/edit.gif border=0></a> <a target=yuzi_frame onclick=checkclick('��ȷ��Ҫɾ������?')  href=manage.asp?menu=deltopic&id="+topicid+"&retopicid="+id+"><img src=images/del.gif border=0></a></td><td valign=bottom><img src=images/posttime.gif> ����ʱ�䣺"+posttime+"��</td><td valign=bottom><img src=images/ip.gif> IP��<a href=manage.asp?menu=lookip&id="+topicid+"&retopicid="+id+">�Ѽ�¼</a></td></tr></table></td></tr></table></td></tr></table>");
}



function ShowForum(ID,topic,newtopic,username,Views,icon,toptopic,locktopic,pollresult,goodtopic,replies,lastname,lasttime)
{
topic = topic.replace(key_word,"<font color=red>"+key_word+"<\/font>");
if (toptopic == 2) {reimage="<img src=images/top.gif>"}
else
if (toptopic == 1) {reimage="<img src=images/f_top.gif>"}
else
if (locktopic== 1) {reimage="<img src=images/f_locked.gif>"}
else
if (pollresult!= '') {reimage="<img src=images/f_poll.gif>"}
else
if ((replies>15) || (Views>150)) {reimage="<img src=images/f_hot.gif>"}
else
if (replies>0) {reimage="<img src=images/f_new.gif>"}
else{reimage="<img src=images/f_norm.gif>"}

if (goodtopic== 1) {reimage2="<img src=images/topicgood.gif>"}
else
if (username == cookieusername) {reimage2="<img src=images/my.gif>"}
else{reimage2=""}

if (replies>0) {reimage3=replies}
else{reimage3="-"}

document.write("<tr height=25><td align=middle width=3% class=a4>"+reimage+"</td>");
document.write("<td width=3% class=a3 align=center>"+reimage2+"</td>");
document.write("<td class=a4 width=45%>&nbsp;<img loaded=no src=images/plus.gif id=followImg"+ID+" style=cursor:hand; onclick=loadThreadFollow("+ID+")> <a target=_blank href=ShowPost.htm?id="+ID+"><img border=0 src=images/brow/"+icon+".gif></a> <a onclick=min_yuzi() target=message href=ShowPost.asp?id="+ID+">"+topic+"</a>");

if (replies > 15) {
var topicpage=""
var tol=replies/15+1;

for (var i=1; i < tol; i++) {
if(i<4 || i>=tol-2){
topicpage=topicpage+"<b><a onclick=min_yuzi() target=message href=ShowPost.asp?id="+ID+"&topage="+ i +">"+ i +"</a></b> ";
}
if (i >= tol-3  && i<tol-2 && i>3){topicpage=topicpage+" ... ";}
}
document.write(" ( <img src=images/multipage.gif> "+topicpage+")");
}

document.write(" "+newtopic+"</td><td align=middle width=9% class=a3><a href=Profile.asp?username="+username+">"+username+"</a></td><td align=middle width=6% class=a4>"+reimage3+"</td><td align=middle width=7% class=a3>"+Views+"</td><td width=27% class=a4>&nbsp;"+lasttime+" by <a href=Profile.asp?username="+lastname+">"+lastname+"</a></td></tr>");
document.write("<tr height=25 style=display:none id=follow"+ID+"><td width=3% class=a4>��</td><td width=3% class=a3>��</td><td id=followTd"+ID+" align=left class=a4 width=94% colspan=5><div onclick=loadThreadFollow("+ID+")><table width=100% cellpadding=10><tr><td width=100%>Loading...</td></tr></table></div></td></tr>");
}


function loadThreadFollow(ino,online){
	var targetImg =eval("followImg" + ino);
	var targetDiv =eval("follow" + ino);
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			targetImg.src="images/minus.gif";
			if (targetImg.loaded=="no"){document.frames["hiddenframe"].location.replace("loading.asp?id="+ino+"&forumid="+online+"");}
		}else{
			targetDiv.style.display="none";
			targetImg.src="images/plus.gif";
		}
}

