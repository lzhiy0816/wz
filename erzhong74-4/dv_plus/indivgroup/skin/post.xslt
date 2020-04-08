<?xml version="1.0" encoding="gb2312"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" >
<xsl:output method="xml" omit-xml-declaration = "yes" indent="yes" version="4.0"/>
	<!--
	Copyright (C) 2004,2005 AspSky.Net. All rights reserved.
	Written by dvbbs.net Lao Mi
	Web: http://www.aspsky.net/,http://www.dvbbs.net/
	Email: eway@aspsky.net
	-->
<xsl:template match="/">
<script language="JavaScript" src="Dv_plus/IndivGroup/js/IndivGroup_Main.js"></script>
<script language="javascript" src="Dv_plus/IndivGroup/js/Dv_Main.js" type="text/javascript"></script>
<script language="javascript" src="Dv_plus/IndivGroup/js/Dv_form.js" type="text/javascript"></script>
<script type="text/javascript" src="Dv_plus/IndivGroup/Edit_Plus/FCKeditor/fckeditor.js"></script>
<xsl:for-each select="/xml/postdata">
	<FORM METHOD="POST" ACTION="{@action}" name="PostForm" onSubmit="return CheckPostForm(this);">
	<input type="hidden" name="RootID" value="{@rootid}"/>
	<input type="hidden" name="PostID" value="{@postid}"/>
	<div class="th"><div style="text-align:left;margin-left:6px;"><xsl:value-of select="@actionname" /></div></div>
	<div class="list">
		<div style="float:left;margin-left:20px;width:10%;border-right:1px solid #559AE4;overflow:hidden;"><b>�û�����</b></div>
		<div style="margin-left:1px;"><input type="text" name="PostUserName" value="{@postusername}"/></div>
	</div>
	<div class="list">
		<div style="float:left;margin-left:20px;width:10%;border-right:1px solid #559AE4;overflow:hidden;"><b>�ꡡ�⣺</b></div>
		<div style="margin-left:1px;"><input type="text" name="Title" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,250);" value="{@topic}"/></div>
	</div>
	<!--
	<div class="list">
		<div style="float:left;margin-left:20px;width:10%;border-right:1px solid #559AE4;overflow:hidden;"><b>�ļ��ϴ���</b><input type="hidden" name="upfilerename" value=""/></div>
		<div style="margin-left:1px;">
			<iframe name="ad" frameborder="0" width="99%" height="24" scrolling="no" src="post_upload.asp?boardid=6"></iframe>
			<iframe name="UPLOADFRAME" frameborder="0" width="80%" height="24" scrolling="no" src="BokeUpload.asp?mode=UploadForm"></iframe>
		</div>
	</div>
	-->
	<div class="list" style="height:375px;">
		<div style="float:left;margin-left:20px;width:10%;height:375px;border-right:1px solid #559AE4;overflow:hidden;"><b>�ڡ��ݣ�</b></div>
		<div style="margin-left:1px;">
			<div style="width:100%;height:348px;"><textarea rows="4" name="PostContent" id="PostContent"><xsl:value-of select="@content"/></textarea></div>
			<div style="margin-left:3px;">UBB��ǩ��<a href="javascript:onclick=PlayUbb('PostContent','FLASH');" alt="��� FLASH ����">[FLASH]</a>
			<a href="javascript:onclick=PlayUbb('PostContent','RM');" alt="��� RM ý���ļ�">[RM]</a>
			<a href="javascript:onclick=PlayUbb('PostContent','MP');" alt="��� MEDIA PLAYER ý���ļ�">[MP]</a>
			<a href="javascript:onclick=PlayUbb('PostContent','QT');" alt="">[QT]</a>
			<a href="javascript:onclick=put_ubb('PostContent','CODE');" alt="Դ������">[CODE]</a></div>
		</div>
	</div>
	<div class="list">
		<div style="margin-left:20px;"><b>����С��ʿ��</b>�� ���⾡���ܼ�����Ҫ�����ݾ�����˵������飬�����ظ��߲�����ȷ����ʱ�Ļظ�</div>
	</div>
	<div class="list" style="background-color:#e4e8ef;"><div style="text-align:center;margin-top:6px;"><input type="submit" name="submit" value="�ύ��Ϣ"/></div></div>
	</FORM>
</xsl:for-each>
<SCRIPT language="JavaScript">
window.onload = function()
{
	var oFCKeditor = new FCKeditor( 'PostContent','100%','350px','<xsl:value-of select="@editmode"/>' ) ;
	oFCKeditor.BasePath	= "Dv_plus/IndivGroup/Edit_Plus/FCKeditor/" ;
	oFCKeditor.Config['SkinPath'] = "skins/office2003/"
	oFCKeditor.ReplaceTextarea() ;
}
//Dv_Form.Set_Select("Best","{@isbest}")
</SCRIPT>
<br/>
</xsl:template>
</xsl:stylesheet>