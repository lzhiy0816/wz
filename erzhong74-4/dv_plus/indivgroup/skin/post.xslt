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
		<div style="float:left;margin-left:20px;width:10%;border-right:1px solid #559AE4;overflow:hidden;"><b>用户名：</b></div>
		<div style="margin-left:1px;"><input type="text" name="PostUserName" value="{@postusername}"/></div>
	</div>
	<div class="list">
		<div style="float:left;margin-left:20px;width:10%;border-right:1px solid #559AE4;overflow:hidden;"><b>标　题：</b></div>
		<div style="margin-left:1px;"><input type="text" name="Title" size="50" onkeyup="javascript:Dv_Form.CheckLength(this,250);" value="{@topic}"/></div>
	</div>
	<!--
	<div class="list">
		<div style="float:left;margin-left:20px;width:10%;border-right:1px solid #559AE4;overflow:hidden;"><b>文件上传：</b><input type="hidden" name="upfilerename" value=""/></div>
		<div style="margin-left:1px;">
			<iframe name="ad" frameborder="0" width="99%" height="24" scrolling="no" src="post_upload.asp?boardid=6"></iframe>
			<iframe name="UPLOADFRAME" frameborder="0" width="80%" height="24" scrolling="no" src="BokeUpload.asp?mode=UploadForm"></iframe>
		</div>
	</div>
	-->
	<div class="list" style="height:375px;">
		<div style="float:left;margin-left:20px;width:10%;height:375px;border-right:1px solid #559AE4;overflow:hidden;"><b>内　容：</b></div>
		<div style="margin-left:1px;">
			<div style="width:100%;height:348px;"><textarea rows="4" name="PostContent" id="PostContent"><xsl:value-of select="@content"/></textarea></div>
			<div style="margin-left:3px;">UBB标签：<a href="javascript:onclick=PlayUbb('PostContent','FLASH');" alt="添加 FLASH 动画">[FLASH]</a>
			<a href="javascript:onclick=PlayUbb('PostContent','RM');" alt="添加 RM 媒体文件">[RM]</a>
			<a href="javascript:onclick=PlayUbb('PostContent','MP');" alt="添加 MEDIA PLAYER 媒体文件">[MP]</a>
			<a href="javascript:onclick=PlayUbb('PostContent','QT');" alt="">[QT]</a>
			<a href="javascript:onclick=put_ubb('PostContent','CODE');" alt="源程序标记">[CODE]</a></div>
		</div>
	</div>
	<div class="list">
		<div style="margin-left:20px;"><b>发贴小贴士：</b>① 标题尽可能简明扼要，内容尽可能说清楚事情，这样回复者才能明确并及时的回复</div>
	</div>
	<div class="list" style="background-color:#e4e8ef;"><div style="text-align:center;margin-top:6px;"><input type="submit" name="submit" value="提交信息"/></div></div>
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