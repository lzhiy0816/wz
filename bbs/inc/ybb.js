
function ybbcode(temp) {

if (badwords!=''){    //����������
var badword= badwords.split ('|'); 
for(i=0;i<badword.length;i=i+1) {
if (badword[i] !=""){temp = temp.replace(badword[i],"****");}
}
}

if (getCookie('ybbcode')!='0'){ //�Ƿ��YBB����
temp = temp.replace(/&amp;/ig,"&");
temp = temp.replace(/  /ig,"��");
temp = temp.replace(/\[b\]/ig,"<b>");
temp = temp.replace(/\[\/b\]/ig,"<\/b>");
temp = temp.replace(/\[i\]/ig,"<i>");
temp = temp.replace(/\[\/i\]/ig,"<\/i>");
temp = temp.replace(/\[u\]/ig,"<u>");
temp = temp.replace(/\[\/u\]/ig,"<\/u>");
temp = temp.replace(/\[strike\]/ig,"<strike>");
temp = temp.replace(/\[\/strike\]/ig,"<\/strike>");
temp = temp.replace(/\[center\]/ig,"<center>");
temp = temp.replace(/\[\/center\]/ig,"<\/center>");
temp = temp.replace(/\[marquee\]/ig,"<marquee>");
temp = temp.replace(/\[\/marquee\]/ig,"<\/marquee>");
temp = temp.replace(/(\[font=)([^.:;`'"=\]]*)(\])/ig,"<FONT face='$2'>");
temp = temp.replace(/\[\/font\]/ig,"<\/FONT>");
temp = temp.replace(/(\[size=)([^.:;`'"=\]]*)(\])/ig,"<FONT size='$2'>");
temp = temp.replace(/\[\/size\]/ig,"<\/FONT>");
temp = temp.replace(/(\[COLOR=)([^.:;`'"=\]]*)(\])/ig,"<FONT COLOR='$2'>");
temp = temp.replace(/\[\/COLOR\]/ig,"<\/FONT>");
temp = temp.replace(/(\[URL\])([^]]*)(\[\/URL\])/ig,"<A TARGET=_blank HREF='$2'>$2</A>");
temp = temp.replace(/(\[URL=)([^]]*)(\])/ig,"<A TARGET=_blank HREF='$2'>");
temp = temp.replace(/\[\/URL\]/ig,"<\/A>");
temp = temp.replace(/(\[EMAIL\])(\S+\@[^]]*)(\[\/EMAIL\])/ig,"<a href=mailto:$2>$2</a>");
temp = temp.replace(/(\[code\])([^]]*)(\[\/code\])/ig,"<BLOCKQUOTE><strong>����</strong>��<HR Size=1>$2<HR SIZE=1><\/BLOCKQUOTE>");
temp = temp.replace(/(\[QUOTE\])(.*)(\[\/QUOTE\])/ig,"<BLOCKQUOTE><strong>����</strong>��<HR Size=1>$2<HR SIZE=1><\/BLOCKQUOTE>");
temp = temp.replace(/\[RM\]([^]]*)\[\/RM\]/ig,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA width=500 height=375><PARAM NAME=SRC VALUE=$1><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true><\/OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=60 width=500><PARAM NAME=SRC VALUE=$1><PARAM NAME=CONTROLS VALUE=ControlPanel,StatusBar><PARAM NAME=CONSOLE VALUE=Clip1><\/OBJECT>");
temp = temp.replace(/(\[MP\])([^]]*)(\[\/MP\])/ig,"<object classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 width=500 height=420><param name=ShowStatusBar value=-1><param name=Filename value=$2><\/object>");
}

if (getCookie('ybbflash')!='0'){  //�Ƿ��[FLASH]����
temp = temp.replace(/(\[FLASH\])([^]]*)(\[\/FLASH\])/ig,"<embed src='$2' width=500 height=375 wmode=transparent>");
}

if (getCookie('ybbimg')!='0'){  //�Ƿ��[IMG]����
temp = temp.replace(/(\[IMG\])([^]]*)(\[\/IMG\])/ig,"<img border=0 src='$2' onmousewheel='return yuzi_img(event,this)' onload='javascript:if(this.width>body.clientHeight)this.width=body.clientHeight'> ");
}

if (getCookie('ybbbrow')!='0'){  //�Ƿ�򿪱������
temp = temp.replace(/(\[em)([0-9]*)(\])/ig,"<IMG border=0 SRC=images/Emotions/$2.gif>");
}

//�Զ�ʶ��URL
temp = temp.replace(/([^]='])(| |<br>)((http|ftp|rtsp|mms):(\/\/|\\\\)[^< ��]+)(| |<br>)/ig,"$1$2<A TARGET=_blank HREF='$3'>$3</A>");

return (temp);
}


function level(experience,membercode,username,moderated)
{
if (membercode=='2'){levelname="�����α�";levelimage="<img src=images/level14.gif>";}
else
if (membercode=='4'){levelname="�� �� Ա";levelimage="<img src=images/level16.gif>";}
else
if (membercode=='5'){levelname="��������";levelimage="<img src=images/level17.gif>";}
else
if(moderated.indexOf("|"+username+"|") != -1 && moderated!=""){levelname="��̳����";levelimage="<img src=images/level15.gif>";}
else
if (experience< 100){levelname="������·";levelimage="<img src=images/level1.gif>";}
else
if (experience< 501){levelname="��������";levelimage="<img src=images/level2.gif>";}
else
if (experience< 1301){levelname="һ��վ��";levelimage="<img src=images/level3.gif>";}
else
if (experience< 2601){levelname="�м�վ��";levelimage="<img src=images/level4.gif>";}
else
if (experience< 4501){levelname="�߼�վ��";levelimage="<img src=images/level5.gif>";}
else
if (experience< 7001){levelname="�� վ ��";levelimage="<img src=images/level6.gif>";}
else
if (experience< 11001){levelname="��ͭ����";levelimage="<img src=images/level7.gif>";}
else
if (experience< 19001){levelname="��������";levelimage="<img src=images/level8.gif>";}
else
if (experience< 30001){levelname="�ƽ���";levelimage="<img src=images/level9.gif>";}
else
if (experience< 45001){levelname="�׽���";levelimage="<img src=images/level10.gif>";}
else
if (experience< 65001){levelname="��ʯ����";levelimage="<img src=images/level11.gif>";}
else
if (experience< 90001){levelname="��վԪ��";levelimage="<img src=images/level12.gif>";}
else
if (experience> 90000){levelname="��������";levelimage="<img src=images/level13.gif>";}
return('');
}

//��ȡCOOKIE
function getCookie (CookieName) { 
var CookieString = document.cookie; 
var CookieSet = CookieString.split (';'); 
var SetSize = CookieSet.length; 
var CookiePieces 
var ReturnValue = ""; 
var x = 0; 
for (x = 0; ((x < SetSize) && (ReturnValue == "")); x++) { 
CookiePieces = CookieSet[x].split ('='); 

if (CookiePieces[0].substring (0,1) == ' ') { 
CookiePieces[0] = CookiePieces[0].substring (1, CookiePieces[0].length); 
}

if (CookiePieces[0] == CookieName) {
ReturnValue = CookiePieces[1];
var value =ReturnValue
}


}
return value;
}
