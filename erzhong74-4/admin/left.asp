<!--#include file="../Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
CheckAdmin(",")
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312" />
<TITLE>myspace</TITLE>
<style type="text/css">
body { margin:0px; background:transparent; overflow:hidden; background:url("skins/images/leftbg.gif"); }
.left_color { text-align:right; }
.left_color a { color: #083772; text-decoration: none; font-size:12px; display:block !important; display:inline; width:175px !important; width:180px; text-align:right; background:url("skins/images/menubg.gif") right no-repeat; height:23px; line-height:23px; padding-right:10px; margin-bottom:2px;}
.left_color a:hover { color: #7B2E00;  background:url("skins/images/menubg_hover.gif") right no-repeat; }
img { float:none; vertical-align:middle; }
#on { background:#fff url("skins/images/menubg_on.gif") right no-repeat; color:#f20; font-weight:bold; }
hr { width:90%; text-align:left; size:0; height:0px; border-top:1px solid #46A0C8;}
</style>
<script type="text/javascript">
<!--
	function disp(n){
		for (var i=0;i<10;i++)
		{
			if (!document.getElementById("left"+i)) return;			
			document.getElementById("left"+i).style.display="none";
		}
		document.getElementById("left"+n).style.display="";
	}
//-->
</script>
</head>
<BODY>


<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" style="padding-top:10px;" class="left_color" id="menubar">
			<div id="left0" style="display:"> 
				<a href="setting.asp" target="frmright">��̳��������</a>
				<a href=CaptchaSet.asp target="frmright">��֤����ʾ����</a>
				<a href="board.asp?action=add" target="frmright">������̳����</a>
				<a href="board.asp" target="frmright">��̳�������</a>
				<a href="user.asp" target="frmright">�û����Ϲ���</a>
				<a href="group.asp" target="frmright">�û���(�ȼ�)����</a>
				<a href="template.asp" target="frmright">���ģ�����</a>
				<a href="admin.asp" target="frmright">����Ա����</a>
				<a href="data.asp?action=BackupData" target="frmright">��̳���ݱ���</a>
				<a href="indivGroup.asp" target="frmright">��̳Ȧ�ӹ���</a>
				<a href="myspace.asp" target="frmright">���˿ռ����</a>
				<a href="plus_adsali.asp" target="frmright">վ��Ӫ��(��������������)</a>
				<a href="forumads.asp" target="frmright">��̳������</a>
				<a href="ReloadForumCache.asp" target="frmright">������̳����</a>
				<a href="log.asp" target="frmright">��̳ϵͳ��־</a>
			</div>

			<div id="left1" style="display:none"> 
				<a href=setting.asp target="frmright">��̳��������</a>
				<a href=CaptchaSet.asp target="frmright">��֤����ʾ����</a>
				<a href=forumads.asp target="frmright">��̳�������</a>
				<a href="plus_adsali.asp" target="frmright"><font color="red">վ��Ӫ��(��������������)</font></a>
				<a href="../announcements.asp?boardid=0&action=AddAnn" target="_blank">��̳�������</a>
				<a href=link.asp?action=add target="frmright">������̳����</a>
				<a href=link.asp target="frmright">������̳����</a>
				<a href="ForumPay.asp" target="frmright">��̳���׹���</a>
				<a href="ForumNewsSetting.asp" target="frmright">��̳��ҳ����</a>
				<a href="badword.asp?reaction=badword" target="frmright">�໰��������</a>
				<a href="badword.asp?reaction=splitreg" target="frmright">ע������ַ�</a>
				<a href="lockip.asp?action=add" target="frmright">IP�����޶�����</a>
				<a href="lockip.asp" target="frmright">IP�����޶�����</a>
			</div>

			<div id="left2" style="display:none"> 
				<a href="board.asp?action=add" target="frmright" alt="">����(����)����</a>
				<a href="board.asp" target="frmright" alt="">����(����)����</a>
				<a href="board.asp?action=permission" target="frmright" alt="">�ְ����û�Ȩ������</a>
				<a href="boardunite.asp" target="frmright" alt="">�ϲ���������</a>
				<a href="update.asp" target="frmright" alt="">�ؼ���̳���ݺ��޸�</a>
			</div>

			<div id="left3" style="display:none">
				<a href="user.asp" target="frmright"  alt="">�û�����(Ȩ��)����</a>
				<a href="group.asp" target="frmright" alt="">�û���(�ȼ�)����</a>
				<a href="wealth.asp" target="frmright" alt="">�û���������</a>
				<a href="message.asp" target="frmright" alt="">�û����Ź���</a>
				<a href="update.asp?action=updateuser" target="frmright" alt="">�ؼ��û���������</a>
				<a href="SendEmail.asp" target="frmright" alt="">�û��ʼ�Ⱥ������</a>
				<a href="admin.asp?action=add" target="frmright" alt="">����Ա����</a>
				<a href="admin.asp" target="frmright" alt="">����Ա����</a>
			</div>

			<div id="left4" style="display:none"> 
				<a href="template.asp" target="frmright" alt="">������ģ���ܹ���</a>
				<a href="Template_RegAndLogout.asp" target="frmright" alt="">ģ��ע����ע��</a>
				<a href="Label.asp" target="frmright" alt="">�Զ����ǩ����</a>
			</div>

			<div id="left5" style="display:none"> 
				<a href="alldel.asp" target="frmright" alt="">����ɾ��</a>
				<a href="alldel.asp?action=moveinfo" target="frmright" alt="">�����ƶ�</a>
				<a href="../recycle.asp" target="frmright" alt="">����վ����</a>
				<a href="postdata.asp?action=Nowused" target="frmright" alt="">��ǰ�������ݱ����� </a>
				<a href="postdata.asp" target="frmright" alt="">���ݱ�������ת�� </a>
			</div>

			<div id="left6" style="display:none"> 
				<a href="data.asp?action=CompressData" target="frmright" alt="">ѹ�����ݿ�</a>
				<a href="data.asp?action=BackupData" target="frmright" alt="">�������ݿ�</a>
				<a href="data.asp?action=RestoreData" target="frmright" alt="">�ָ����ݿ�</a>
				<a href="address.asp?action=add" target="frmright" alt="">��̳IP������</a>
				<a href="address.asp" target="frmright" alt="">��̳IP����� </a>
			</div>

			<div id="left7" style="display:none"> 
				<a href="upUserface.asp" target="frmright" alt="">�ϴ�ͷ�����</a>
				<a href="uploadlist.asp" target="frmright" alt="">�ϴ��ļ�����</a>
				<a href="bbsface.asp?Stype=3" target="frmright" alt="">ע��ͷ�����</a>
				<a href="bbsface.asp?Stype=2" target="frmright" alt="">�����������</a>
				<a href="bbsface.asp?Stype=1" target="frmright" alt="">�����������</a>
			</div>

			<div id="left8" style="display:none"> 
				<a href="plus.asp" target="frmright" alt="">��̳�˵�����</a>				
				<a href="indivGroup.asp" target="frmright" alt="">����Ȧ�ӹ���</a>
				<a href="../bokeadmin.asp" target="frmright" alt="">��̳���͹���</a>
				<a href="myspace.asp" target="frmright" alt="">���˿ռ����</a>
				<a href="plus_Tools_Info.asp?action=Setting" target="frmright" alt="">������������</a>
				<a href="plus_Tools_Info.asp?action=List" target="frmright" alt="">������������</a>
				<a href="plus_Tools_User.asp" target="frmright" alt="">�û����߹���</a>
				<a href="plus_Tools_User.asp?action=paylist" target="frmright" alt="">������Ϣ����</a>
				<a href="MoneyLog.asp" target="frmright" alt="">����������־</a>
				<a href="plus_Tools_Magicface.asp" target="frmright" alt="">ħ����������</a>
				<a href="plus_cnzz_wss.asp" target='frmright'>����ͳ��</a>
				<a href="plus_ccvideo.asp" target="frmright" alt="">CC��Ƶ���</a>
				<a href="plus_qcomic.asp" target='frmright'>��ͼ��������</a>				
			</div>

			<div id="left9" style="display:none"> 
				<a href="data.asp?action=SpaceSize" target="frmright" alt="">ϵͳ��Ϣ���</a>
				<a href="log.asp" target="frmright" alt="">��̳ϵͳ��־</a>
				<a href="help.asp" target="frmright" alt="">��̳��������</a>
				<a href="ReloadForumCache.asp" target="frmright" alt="">������̳����</a>
			</div>
	</td>
 </tr>
</table>
</body>
</html>