function presskey(eventobject){
if(event.ctrlKey && window.event.keyCode==13){
ValidateForm()
this.document.form.submit();}
}

function HighlightAll(theField) {
var tempval=eval("document."+theField)
tempval.focus()
tempval.select()
therange=tempval.createTextRange()
therange.execCommand("Copy")
}

function replac(){
txt2=prompt("��������","");
if (txt2 != "") {txt=prompt("�滻Ϊ��",txt2)}else {replac()}
var Otext = txt2; var Itext = txt; document.form.content.value = eval('form.content.value.replace(/'+Otext+'/'+'g'+',"'+Itext+'")')
}

function DoTitle(addTitle) {
var revisedTitle;var currentTitle = document.form.topic.value;revisedTitle = addTitle+currentTitle;document.form.topic.value=revisedTitle;document.form.topic.focus();
}


function ValidateForm(){
MessageLength =document.form.content.value.length;
if(MessageLength<2){alert("�������ݲ���С��2���ַ���");return false;}
sending.style.visibility="visible";
}

function CheckLength(){
var MessageMax="50000";
MessageLength=document.form.content.value.length;
alert("����ַ�Ϊ "+MessageMax+ " �ֽ�\n������������ "+MessageLength+" �ֽ�");
}



function emoticon(theSmilie){
document.form.content.value += theSmilie + ' ';
document.form.content.focus();
}


function showadv(){
if (document.form.advshow.checked == true) {
adv.style.display = "";
}else{
adv.style.display = "none";
}
}