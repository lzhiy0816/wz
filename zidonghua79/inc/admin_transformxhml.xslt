<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
<xsl:variable name="pagesize" select="50"/>
<xsl:template match="/">
<script language="Javascript" src="images/post/reply.js" type="text/javascript"></script>
<iframe style="border:0px;width:0px;height:0px;"  src="" name="hid1"></iframe>
<div class="mainbar0">
<fieldset style="border : 1px dotted #ccc;text-align : left;line-height:22px;">
<legend>
<b>����˵��</b>
</legend>
<ol style="margin-top:0px;margin-bottom:0px;list-style-position : inside;">
<li style="padding:0px 5px;">���������԰�����������ת���ɹ淶��XHTML��ʽ��������̳���ġ�</li>
<li style="padding:0px 5px;">��ѡ����Ҫת���ı���д��ʼת����������ţ����ύ��</li>
<li style="padding:0px 5px;">ע�⣺������Ҫ�ϳ���ʱ�䣬����ռ�ô����ڴ棬�����ڱ���ִ�С�ִ�й����в���ˢ��ҳ���رմ��ڣ������ĵȴ���</li>
<li style="padding:0px 5px;">�������л�����������Ӹ�ʽ����Ϊ�����������������ʧ��ǿ�ҽ������ȱ��ݺ��������ݿ⡣�����ʹ�õ���Access���ݿ⣬������ִ�б�����֮ǰ��֮�󣬶�Ҫִ�����ݿ���޸���ѹ����</li>
</ol>
</fieldset>
</div>
<div class="th" style="text-indent:10px;">��������</div>
<form id="Dvform1" name="Dvform1" action="Admin_transformxhml.asp?action=transform" method="post" target="hid1">
<div class="mainbar3" style="text-align : left;line-height:25px;height:25px;text-indent:10px;">
��ѡ���޸������ݱ�
<select name="tableid" id="tableid">
<xsl:for-each select="xml/posttable">
<option value="{@id}"><xsl:if test="@id= 1"><xsl:attribute name="selected">selected</xsl:attribute></xsl:if><xsl:value-of select="@tabletype" />[��¼��<xsl:value-of select="@count" />]</option>
</xsl:for-each>
</select>
��ʼ��ţ�<input type="text" name="sid" value="1" />
ÿ���޸�����<input type="text" name="transcount" value="1000" />
<input type="submit" value="�� ʼ" name="submit"/>
</div>
</form>
<div class="mainbar0">
<fieldset style="border : 1px dotted #ccc;text-align : left;line-height:22px;">
<legend>
<b>ת�����</b>
</legend>
<div style="text-indent:24px;" id="title1">��δ��ʼ</div>
<div style="text-indent:24px;" id="title2"></div>
</fieldset>
<hr />
<iframe class="Dvbbs_Reply" id="Dvbbs_Composition" marginheight="5" marginwidth="5" width="0" height="0" ></iframe>
</div>
<form id="Dvform" name="Dvform" action="?action=dotransform" method="post" target="hid1">
<input type="hidden" id="Body" name="Body" value="" />
<input type="hidden" id="id" name="id" value="1" />
<input type="hidden" id="tableid" name="tableid" value="1" />
</form>
<script language="Javascript" type="text/javascript">
var Dvbbs_bIsIE5=document.all;
var Dvbbs_Mode = 3;
if (Dvbbs_bIsIE5)
{var IframeID=frames["Dvbbs_Composition"];}
else
{
var IframeID=document.getElementById("Dvbbs_Composition").contentWindow;
}
Dvbbs_InitDocument('Body','gb2312');
function copyto(text){
IframeID.document.body.innerHTML=text;
Dvbbs_CopyData('Body');
document.Dvform.submit()
}
</script>
</xsl:template>

</xsl:stylesheet>
