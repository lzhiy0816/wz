/*
POPWAN API SCRIPT CODE

*/
/*
ʹ�������Ӧ����ҳ�߶�
<iframe id="pw_frame" src="@url" scrolling="no" frameborder="0"></iframe>
ʹ�ã�window.setInterval("reinitIframe('pw_frame')", 200);
*/
function reinitIframe(FID){
	var iframe = document.getElementById(FID);
	try{
		var bHeight = iframe.contentWindow.document.body.scrollHeight;
		var dHeight = iframe.contentWindow.document.documentElement.scrollHeight;
		var height = Math.max(bHeight, dHeight);
		iframe.height =  height;
	}catch (ex){}
}


function checkHeight(FID) {
	var iframe = document.getElementById(FID);
	var bHeight = iframe.contentWindow.document.body.scrollHeight;
	var dHeight = iframe.contentWindow.document.documentElement.scrollHeight;
	alert("bHeight:" + bHeight + ", dHeight:" + dHeight);
}

/*�����¼�*/
function copyText(obj)
{
	if (!obj)
	{
		return false;
	}
    if( obj.type !="hidden" )
    {
        obj.focus();
    }
    obj.select(); 
    copyToClipboard(obj.value);
    alert("�����ɹ���");
}

function copyToClipboard(txt) {  
    if(window.clipboardData)  
    {  
        window.clipboardData.clearData();  
        window.clipboardData.setData("Text", txt);  
    }  
    else if(navigator.userAgent.indexOf("Opera") != -1)  
    {  
        window.location = txt;  
    }  
    else if (window.netscape)  
    {  
        try {  
            netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");  
        }  
        catch (e)  
        {  
            alert("��ʹ�õ��������֧�ִ˸��ƹ��ܣ���ʹ��ctrl+c����������Ҽ�����");  
        }  
        var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);  
        if (!clip)  
            return;  
        var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);  
        if (!trans)  
            return;  
        trans.addDataFlavor('text/unicode');  
        var str = new Object();  
        var len = new Object();  
        var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);  
        var copytext = txt;  
        str.data = copytext;  
        trans.setTransferData("text/unicode",str,copytext.length*2);  
        var clipid = Components.interfaces.nsIClipboard;  
        if (!clip)  
            return false;  
        clip.setData(trans,null,clipid.kGlobalClipboard);  
    }  
    return true;  
} 