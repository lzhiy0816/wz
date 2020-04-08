function ybbsize(theSmilie){
var text=prompt("请输入文字", "");
if(text){
document.form.content.value += '[size=' + theSmilie + ']'+ text + '[/size]';
}
}


function ybbfont(theSmilie){
var text=prompt("请输入文字", "");
if(text){
document.form.content.value += '[font=' + theSmilie + ']'+ text + '[/font]';
}
}


function YBBurl() {
var enterURL   = prompt("请输入主页地址", "http://");
if (enterURL) {
var ToAdd = "[URL]"+enterURL+"[/URL]";
document.form.content.value+=ToAdd;
}
}


function YBBimage() {
var enterURL   = prompt("请输入图片地址", "http://");
if (enterURL) {
var ToAdd = "[IMG]"+enterURL+"[/IMG]";
document.form.content.value+=ToAdd;
}
}


function YBBemail() {
var emailAddress = prompt("请输入邮件地址","");
if (emailAddress) {
var ToAdd = "[EMAIL]"+emailAddress+"[/EMAIL]";
document.form.content.value+=ToAdd;
}
}

function YBBflash() {
var enterURL   = prompt("请输入FLASH地址", "http://");
if (enterURL) {
var ToAdd = "[FLASH]"+enterURL+"[/FLASH]";
document.form.content.value+=ToAdd;
}
}

function YBBrm() {
var enterURL   = prompt("请输入RealPlayer文件地址", "");
if (enterURL) {
var ToAdd = "[RM]"+enterURL+"[/RM]";
document.form.content.value+=ToAdd;
}
}

function YBBmp() {
var enterURL   = prompt("请输入Media Player文件地址", "");
if (enterURL) {
var ToAdd = "[MP]"+enterURL+"[/MP]";
document.form.content.value+=ToAdd;
}
}


function YBBbold() {
document.form.content.value="[B]"+document.form.content.value+"[/B]";
}


function YBBitalic() {
document.form.content.value="[I]"+document.form.content.value+"[/I]";
}


function YBBunder() {
document.form.content.value="[U]"+document.form.content.value+"[/U]";
}

function YBBcenter() {
document.form.content.value="[CENTER]"+document.form.content.value+"[/CENTER]";
}

function YBBquote() {
document.form.content.value="[QUOTE]"+document.form.content.value+"[/QUOTE]";
}



function MM_findObj(n, d) { 
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i>d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function insertTag(MyString)
{
fontbegin='[color=' + MyString + ']'
fontend='[/color]'
document.form.content.value=fontbegin+document.form.content.value+fontend;
document.form.content.focus();
}


var base_hexa = "0123456789ABCDEF";
function dec2Hexa(number)
{
   return base_hexa.charAt(Math.floor(number / 16)) + base_hexa.charAt(number % 16);
}

function RGB2Hexa(TR,TG,TB)
{
  return "#" + dec2Hexa(TR) + dec2Hexa(TG) + dec2Hexa(TB);
}

function lightCase(MyObject,objName)
{
	MM_findObj(objName).bgColor = MyObject.bgColor;
}

col = new Array;
col[0] = new Array(255,0,255,-1,255,-1);
col[1] = new Array(255,0,0,1,0,0);
col[2] = new Array(255,-1,255,0,0,0);
col[3] = new Array(0,0,255,0,0,1);
col[4] = new Array(0,0,255,-1,255,0);
col[5] = new Array(0,1,0,0,255,0);
col[6] = new Array(255,-1,0,0,255,-1);

function rgb(pas,w,h){
	for (j=0;j<6+1;j++){
		for (i=0;i<pas+1;i++){
			r = Math.floor(col[j][0]+col[j][1]*i*(255)/pas);
			g = Math.floor(col[j][2]+col[j][3]*i*(255)/pas);
			b = Math.floor(col[j][4]+col[j][5]*i*(255)/pas);
		  codehex = r + '' + g + '' + b;
		  document.write('<td bgColor=\"' + RGB2Hexa(r,g,b) + '\"onmouseover="lightCase(this,\'ColorUsed\')" onClick=\"insertTag(this.bgColor)\" width=\"'+w+'\" height=\"'+h+'\"><IMG height='+h+' width='+w+' border=0 title=\"字体颜色：'+RGB2Hexa(r,g,b)+'\"></TD>\n');
		}
	}
}