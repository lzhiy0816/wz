<!-- #include file="setup.asp" --><%
select case Request("menu")

case ""
index

case "guizhe"
guizhe

case "level"
level

case "zhuce"
zhuce

case "daima"
daima

case "YBB"
YBB

case "cook"
cook

case "bz"
bz

case "cz"
cz

case "edit"
edit

case "gb"
gb

end select

sub index
top
%><body bgcolor="#FFFFFF" text="#000000" background="images/lzybg01.jpg">
</body>
<title>��������</title>
<table width=97% align="center" border="0">
  <tr>
    <td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="����ҳ"></a></td>
    <td vAlign="center" align="top">��<img src="images/closedfold.gif" border="0" />��<a href=main.asp><%=clubname%></a><br />
    ��<img src="images/bar.gif" border="0" /><img src="images/openfold.gif" border="0" />����������</td>
    </tr>
  </table>


    <center>
    <br>

<SCRIPT>valigntop()</SCRIPT>
    <table  cellspacing="1" cellPadding="4" width="97%" border="0" class=a2>
      <tr>
        <td vAlign="top" width="193" class=a3><br>
        �����˵�<br>
        <a target="I2" href="help.asp?menu=guizhe">������������</a><br />
        <a target="I2" href="help.asp?menu=level">�ȼ�����</a><br />
        <a target="I2" href="help.asp?menu=zhuce">���Ƿ����ע��</font></a><br />
        <a target="I2" href="help.asp?menu=daima">�Ƿ���Լ���HTML���� 
        </font></a><br />
        <a target="I2" href="help.asp?menu=YBB">YBB���� </font></a>
        <br />
        <a target="I2" href="help.asp?menu=bz">ʲô�ǰ���</font></a><br />
        <a target="I2" href="help.asp?menu=cook">ʹ��COOKIE��</font></a><br />
        <a target="I2" href="help.asp?menu=edit">�����޸��ҵ�ע����Ϣ</font></a><br />
        <a target="I2" href="help.asp?menu=gb">��ѶϢ</font></a><br />
        <a target="I2" href="help.asp?menu=cz">�ܽ��в����� </font></a>
        <br />
        </td>
        <td vAlign="top" width="541" class=a4>
        <p>
        <iframe name="I2" width="100%" height="250" marginwidth="1" marginheight="1" border="0" frameborder="0">
        �������֧��Ƕ��ʽ��ܣ�������Ϊ����ʾǶ��ʽ��ܡ�</iframe></p>
        </td>
        </tr>
        </table>
<SCRIPT>valignbottom()</SCRIPT>
        <%htmlend
end sub





sub level%>
        <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
        <link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
        <script src="inc/ybb.js"></script>
        <table align="left" border="0" cellpadding="3" cellspacing="1" class=a2 width="100%">
          <tr>
            <td colspan="3" class=a1 width="512">
            <p align="center"><a><b>�� ��̳�ȼ����</b></a></p>
            </td>
          </tr>
          <tr bgcolor="A09CFF">
            <td bgcolor="FFFFFF" width="106" align="center">�ȼ�����</td>
            <td bgcolor="FFFFFF" width="237">
            <p align="center">����ֵ</p>
            </td>
            <td bgcolor="FFFFFF" width="155">�Ǽ�</td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF">
            
           
            
            <font face="����"> <script>document.write(""+level(0,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">0 - 100</td>
            <td width="155" bgcolor="FFFFFF">
            <p><font face="����">�Ǽ�Ϊ </font>
<img border="0" src="images/level1.gif" />            </p>
            </td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(101,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">101 - 500</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level2.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(501,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">501 - 1300</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level3.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(1301,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">1301 - 2600</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level4.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(2601,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">2601 - 4500</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level5.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(4501,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">4501 - 7000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level6.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(7001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">7001 - 11000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level7.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(11001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">11001 - 19000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level8.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(19001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">19001 - 30000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level9.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(30001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">30001 - 45000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level10.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(45001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">45001 - 65000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level11.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(65001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">65001 - 90000</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level12.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����"><script>document.write(""+level(90001,1,"","")+levelname+"");</script>��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">90001 - ����</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level13.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����">�����α���</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level14.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����">�� ����</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level15.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����">�� �� Ա��</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level16.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="����">����������</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="����">�Ǽ�Ϊ </font>
            <img border="0" src="images/level17.gif" /></td>
          </tr>
        </table>
        <%end sub
sub guizhe%>
        <meta http-equiv="Content-Type" content="text/html; charset=gb2312"><link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
        <table align="left" border="0" cellpadding="3" cellspacing="1" class=a2 width="100%" height="1">
          <tr>
            <td class=a1 width="512" height="15">
            <p align="center"><font style="font-size:9pt"><a><b>
            <font color="FFFFFF">��&nbsp; ������������</font></b></a></font></p>
            </td>
          </tr>
          <tr bgcolor="A09CFF">
            <td bgcolor="FFFFFF" width="498" height="1">
            <p align="left">1������ֵ�����ԴӸ������������󣬲��ޡ�<br />
            ����&nbsp; ���ֵ�����ԴӸ������������󣬲��ޡ�<br />
            ����&nbsp; ����ֵ��<font color="FF0000">0--100</font></p>
            </td>
            </tr>
            <tr bgcolor="A09CFF">
              <td bgcolor="FFFFFF" width="498" height="67">2����һ�ε�¼��������ʼ����ֵΪ0���Ժ��ۼ�<br />
              ����&nbsp; ��һ�ε�¼��������ʼ���ֵΪ0���Ժ��ۼ�<br />
              ����&nbsp; ��һ�ε�¼����������ֵ��ΪΪ100���Ժ��ۼ�</td>
              </tr>
              <tr bgcolor="A09CFF">
                <td bgcolor="FFFFFF" width="498" height="1">3��д���£�<br />
                ����&nbsp; ����ֵ: <font color="FF0000">+5</font><br />
                ����&nbsp; ���ֵ: <font color="FF0000">+5</font><br />
                ����&nbsp; ����ֵ: <font color="FF0000">-5</font> �������α�������������Ա����������������</td>
                </tr>
                <tr bgcolor="A09CFF">
                  <td bgcolor="FFFFFF" width="498" height="1">4�������£�<br />
                  ����&nbsp; �ظ��߾���ֵ: <font color="FF0000">+2</font><br />
                  ����&nbsp; �ظ��߽��ֵ: <font color="FF0000">+2</font><br />
                  ����&nbsp; �ظ�������ֵ: <font color="FF0000">-2</font> �������α�������������Ա����������������</td>
                  </tr>
                      <tr bgcolor="A09CFF">
                        <td bgcolor="FFFFFF" width="498" height="1">5����Чͣ��ʱ�䣨��Ҫ��¼��������<br />
                      ����&nbsp; ÿ��Чʱ��<font color="FF0000"> 10 </font>���ӣ� ����ֵ:
                      <font color="FF0000">+1</font> ����ֵ��<font color="FF0000">+10</font></td>
                      </tr>
                      <tr bgcolor="A09CFF">
                        <td bgcolor="FFFFFF" width="498" height="1">6������ֵ &lt;
                        <font color="FF0000">1</font> ��������ֵ &lt;
                        <font color="FF0000">5</font> �ߣ������ܷ������¡�</td>
                        </tr>
                        </table>
                        <%end sub
sub zhuce%>
                        
                        <meta http-equiv="Content-Type" content="text/html; charset=gb2312"><link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                        <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2>
                          <tr>
                            <td class=a1>
                            <div align="center">
                              <font style="font-size:9pt"><b>
                              <a name="register1">ע��</a></b></font></div>
                            </td>
                          </tr>
                          <tr>
                            <td bgcolor="FFFFFF"><font style="font-size:9pt">�ǵ�������ע������ڷ����������£� ��ֻ����ʽע�ᣬ������ӵ����ʾ�Լ��������ϵȹ��ܣ�</font></td>
                          </tr>
                        </table>
                        <%end sub
sub daima%>
                        <meta http-equiv="Content-Type" content="text/html; charset=gb2312"><link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                        <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2 style="border-collapse: collapse">
                          <tr>
                            <td class=a1>
                            <div align="center">
                              <a name="html"><b>
                              <font style="font-size:9pt">ʹ��</font><font style="font-size:9pt" face="Verdana, Arial">HTML</font><font style="font-size:9pt">�������̳ר�ô���</font></b></a></div>
                            </td>
                          </tr>
                          <tr>
                            <td bgcolor="FFFFFF"><font style="font-size:9pt">����������������Ϣ��ʹ��</font><font style="font-size:9pt" face="Verdana, Arial">HTML</font><font style="font-size:9pt">���뵫����ʹ����̳ר�ô���<font face="Verdana, Arial">YBB</font>����</font><font style="font-size:9pt" face="Verdana, Arial">.&nbsp;
                            </font><font style="font-size:9pt">��̳ר�ô���������</font><font style="font-size:9pt" face="Verdana, Arial">HTML</font><font style="font-size:9pt">����</font><font style="font-size:9pt" face="Verdana, Arial">,
                            </font><font style="font-size:9pt">��ֻ�ṩ�����������ܺ�һЩר�ù���</font><font style="font-size:9pt" face="Verdana, Arial">,
                            </font><font style="font-size:9pt">��</font><font style="font-size:9pt" face="Verdana, Arial">hyperlinking, 
                            image </font><font style="font-size:9pt">��ʾ</font><font style="font-size:9pt" face="Verdana, Arial">,
                            </font><font style="font-size:9pt">�����б����</font><font style="font-size:9pt" face="Verdana, Arial">.
                            </font><font style="font-size:9pt">
                            <a target="I2" href="help.asp?menu=YBB">����</a>����������̳ר�ô����б�</font><font style="font-size:9pt" face="Verdana, Arial">.
                            </font></td>
                          </tr>
                        </table>
                        <%end sub
sub YBB%>
                        <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
                        <link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                        <table border="0" cellspacing="1" cellpadding="3" class=a2 height="1215">
                          <tr>
                            <td class=a1 height="13">
                            <b>ʲô��YBB(Yuzi 
                            Bulletin Board)���룿</font>
                            </font></b></td>
                          </tr>
                          <tr>
                            <td height="58" bgcolor="F7F7F7">
                            <p><font face="Verdana, Arial">YBB</font><font style="font-size:9pt">������</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">��һ�����֡���Ҳ���Ѿ���������Ϥ�ˡ�һ������£������������</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">��Ҳ�Ϳ���ʹ��</font><font face="Verdana, Arial" style="font-size:9pt">YBB</font><font style="font-size:9pt">���롣��ʹ������������������ʹ��</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">����Ҳ����ʹ��</font><font face="Verdana, Arial" style="font-size:9pt">YBB</font><font style="font-size:9pt">���롣 
                            ����Ҫ��ʹ�õı�����٣���ʹ����ʹ��</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">��������Ҳ��ʹ��</font><font face="Verdana, Arial" style="font-size:9pt">YBB</font><font style="font-size:9pt">���룬��Ϊ�������Ŀ����Դ���С�ˡ�</font></p>
                            </td>
                          </tr>
                          <tr>
                            <td class=a1 height="16">
                            <b>

                            URL��������</b></td>
                          </tr>
                          <tr>
                            <td height="93" bgcolor="FFFFFF">��������Ϣ����볬�����ӣ�ֻҪ�����з�ʽ����Ϳ�����</font><font face="Verdana, Arial">(YBB</font>������<font color="ff0000">����</font></font><font face="Verdana, Arial">).
                            </font><br />
                            </font></font><font face="Verdana, Arial">
                            <br />
                            </font><font face="Verdana, Arial">
                            <font color="ff0000">[url]</font>http://www.yuzi.net<font color="ff0000">[/url]</font>
                            <br />
                            </font></font>���� <br />
                            <font face="Verdana, Arial"><font color="ff0000">[url=</font>http://www.yuzi.net]Yuzi</font>������<font color="ff0000">[/url]</font> <br />
                            </font></font><font><br />
                            ���������룬</font><font face="Verdana, Arial">YBB</font>������Զ���</font><font face="Verdana, Arial">URL</font>�������ӣ�����֤���û�����µĴ���ʱ��������Ǵ��ŵġ�ע��</font><font face="Verdana, Arial">URL</font>��</font><font face="Verdana, Arial">&quot;http://&quot;</font>��һ����������ġ�</font></font></td>
                            </tr>
                            <tr>
                              <td class=a1 height="13">
                              <b>�����ʼ�����</font></b></td>
                            </tr>
                            <tr>
                              <td height="39" bgcolor="FFFFFF">��������Ϣ���������ʼ��ĳ������ӣ�ֻҪ������������Ϳ�����(YBB�����Ǻ���)<br />
                              </font></font>
                              <font color="ff0000"><br />
                              <font color="FF0000">[email]</font></font><font color="000000"><a href="mailto:huangzhiyu@yuzi.net">huangzhiyu@yuzi.net</a></font><font color="ff0000">[/email]</font><br />
                              </font></font>
                              <font color="ff0000"><br />
                              </font>���������룬YBB�����Ե����ʼ��Զ��������ӡ� </font></td>
                              </tr>
                              <tr>
                                <td class=a1 height="13">
                                <b>������б��</font></b></td>
                              </tr>
                              <tr>
                                <td height="74" bgcolor="FFFFFF"><font>������ʹ�� 
                                </font><font face="Verdana, Arial">[b] 
                                [/b] </font>�� </font>
                                <font face="Verdana, Arial">[i] [/i]
                                </font>��Щ��־�Դﵽ��������ʹ�ô����б���Ч��<font face="Verdana, Arial">.</font>
                                </font></font><br />
                                </font></font><br />
                                ����<font face="Verdana, Arial">,
                                <font color="ff0000">[b]</font></font></font><b>����Ա</font></font></b><font color="FF0000">[/b]</font><font face="Verdana, Arial"><br />
                                </font>
                                ����</span></font><font face="Verdana, Arial">,
                                <font color="ff0000">[i]</font></font><i>����Ա</font></i><font color="ff0000">[/i]</font><font face="Verdana, Arial"><br />
                                </font>����,<font color="ff0000">[u]</font><u>����Ա</u><font color="ff0000">[/u]</font></font></span>
                                </font>
                                <font face="Verdana, Arial"><br />
                                </font>����,<font color="ff0000">[strike]</font><strike>����Ա</strike></font><font color="ff0000">[/strike]</font></span></td>
                                </tr>
                                <tr>
                                  <td class=a1 height="13">
                                  <b>�ƶ�����
                                  </font></b></td>
                                </tr>
                                <tr>
                                  <td height="42" bgcolor="FFFFFF">��������Ϣ������ƶ����֣�ֻҪ����������Ϳ�����</font><font face="Verdana, Arial">(YBB</font>������<font color="ff0000">����</font></font><font face="Verdana, Arial">).</font>
                                  </font><font color="ff0000">
                                  <br />
                                  </font><br />
                                  </font></font>
                                  <font color="ff0000">[marquee]</font>�ƶ�����</font><font color="ff0000">[/marquee]</font><br />
                                  </font></font>
                                  <marquee>�ƶ�����</marquee></font></td>
                                  </tr>
                                  <tr>
                                    <td class=a1 height="13">
                                    <b>����������Ϣ
                                    </font></b></td>
                                  </tr>
                                  <tr>
                                    <td height="58" bgcolor="FFFFFF">����һЩ�˵����ӣ�ֻҪ���кͿ���Ȼ����������Ϳ�����</font><font face="Verdana, Arial">(YBB</font>������ 
                                    </font>
                                    <font color="ff0000">����</font><font face="Verdana, Arial" color="ff0000">)</font><font face="Verdana, Arial">).</font> <br />
                                    </font><font color="ff0000"><br />
                                    [QUOTE]</font>�������Ĺ�����Ϊ����ʲô......<br />
                                    ������Ϊ���Ĺ�����ʲô��</font><font color="ff0000">[/QUOTE]</font>
                                    <br />
                                    <br />
                                    �������У�</font><font face="Verdana, Arial">YBB</font></font>������Զ����������õ����֡� 
                                    </font></td>
                                    </tr>
                                      <tr>
                                        <td height="16" class=a1>
                                        <b>��ɫ����</font></b></td>
                                      </tr>
                                      <tr>
                                        <td height="169" bgcolor="FFFFFF">
                                        <p>
                                        <font color="ff0000" face="Verdana, Arial">[color=red]</font>��ɫ����<font color="ff0000">[/color]</font><br />
                                        ���ϵ�&quot;red&quot;�����ǣ�</font></p>
                                        <div align="center">
                                          <center>
                                          <table border="0" width="49%">
                                            <tr>
                                              <td height=25 width="31%" align="center" bgcolor="F7F7F7">
                                              Black </td>
                                              <td height=25 width="35%" align="center" bgcolor="F7F7F7">
                                              #000000</td>
                                              <td height=25 width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ </td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Silver </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #C0C0C0</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Gray </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #808080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Pink </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FFC8CB</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              �ۺ�</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Maroon </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #800000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Red </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FF0000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Purple </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #800080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Fuchsia </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FF00FF</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              �Ϻ�</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Green</td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #008000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Lime </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #00FF00</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ����</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Olive </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #808000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ���</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Yellow </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FFFF00</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Navy </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #000080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ����</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Blue </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #0000FF</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Teal </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #008080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Aqua </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #00FFFF</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ǳ��</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Orange</td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FFA500</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Brown</td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #A52A2A</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              ��ɫ</td>
                                            </tr>
                                          </table>
                                          </center>
                                        </div>
                                        </td>
                                        </tr>
                                          <tr>
                                            <td height="14" class=a1><b>�ر�ע��</b></td>
                                          </tr>
                                          <tr>
                                            <td height="76" bgcolor="FFFFFF">
                                            <p>��������ͬʱʹ��</font><font face="Verdana, Arial">HTML</font>��</font><font face="Verdana, Arial">YBB</font>�����ͬһ�ֹ��ܡ�����ע��</font><font face="Verdana, Arial">YBB</font>���벻�Դ�Сд���С�������������</font><font color="ff0000">[URL]</font>
                                            �� </font>
                                            <font color="ff0000">[url]</font>
                                            </font>
                                            <font color="800000">
                                            <br />
                                            ����ȷ��</font><font color="800000" face="Verdana, Arial">YBB</font><font color="800000">����ʹ�ã�</font><font face="Verdana, Arial"><font color="ff0000"><br />
                                            [url]</font> www.yuzi.net
                                            <font color="ff0000">[/url]</font>
                                            </font>��Ҫ�����ź������������֮���пո�</font><font face="Verdana, Arial"><font color="ff0000"><br />
                                            [email]</font>huangzhiyu@yuzi.net<font color="ff0000">[email]</font>
                                            </font>�ڽ���ʱ����Ҫ�����������ڼ���б��</font><font color="ff0000">[/email]</font>
                                            </font></p>
                                            </td>
                                            </tr>
                                            </table>
                                            <%end sub
sub bz%>
                                            <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
                                            <link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2 style="border-collapse: collapse">
                                              <tr>
                                                <td class=a1>
                                                <div align="center">
                                                  <font style="font-size:9pt">
                                                  <b>����</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                ���������������Ĺ���Ա��������ɾ������������κ����ӡ��������ĳ���������κ����ʣ���ֱ�����������ϵ��</font>
                                                </td>
                                              </tr>
                                            </table>
                                            <%end sub
sub cook%>
                                            <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
                                            <link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2 style="border-collapse: collapse">
                                              <tr>
                                                <td class=a1>
                                                <div align="center">
                                                  <font style="font-size:9pt">
                                                  <b>COOKIE</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                ����̳ʹ����cookie���������Դ洢�����û����Ϳ��cookie��Ϣ�洢������������cookie����������ٻ�����������ã�����ʹ���ܸ��ӷ����ʹ�ñ���̳ϵͳ����������������֧��cookie���������������cookie���ܹرգ���Щʡʱ�Ĺ��ܽ��޷�������</font>
                                                </td>
                                              </tr>
                                            </table>
                                            <%end sub
sub edit%>
                                            <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
                                            <link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2 style="border-collapse: collapse">
                                              <tr>
                                                <td class=a1>
                                                <div align="center">
                                                  <font style="font-size:9pt">
                                                  <b>����������Ϣ</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                ������ͨ��ҳ���Ϸ���</font><font style="font-size:9pt" face="Verdana, Arial">&quot;�������&quot;,
                                                </font>
                                                <font style="font-size:9pt">�ܷ���ĸ�������ע����Ϣ</font><font style="font-size:9pt" face="Verdana, Arial">.
                                                </font>
                                                <font style="font-size:9pt">���������û���������</font><font style="font-size:9pt" face="Verdana, Arial">,
                                                </font>
                                                <font style="font-size:9pt">�������Բ鿴���޸�����ǰ����ע�����Ϣ</font><font style="font-size:9pt" face="Verdana, Arial">,
                                                </font>
                                                <font style="font-size:9pt">��Ȼ���������û���</font><font style="font-size:9pt" face="Verdana, Arial">.
                                                </font></td>
                                              </tr>
                                            </table>
                                            <%end sub
sub gb%>
                                            <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
                                            <link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2 style="border-collapse: collapse">
                                              <tr>
                                                <td class=a1>
                                                <div align="center">
                                                  <font style="font-size:9pt">
                                                  <b>��ѶϢ</b>
                                                  </font>
                                                </div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td height="29" bgcolor="FFFFFF">
                                                <font style="font-size:9pt">�������ڵ�ʱ�򣬱��˿��Է���ѶϢ�����������û��ǲ��ܿ�������ѶϢ��ֻ�����ܹ������Լ���ѶϢ����ÿ���˶����Ը������ԣ�</font></td>
                                              </tr>
                                            </table>
                                            <%end sub
sub cz%>
                                            <meta http-equiv="Content-Type" content="text/html; charset=gb2312"><link href="images/skins/<%=Request.Cookies("skins")%>/bbs.css" rel="stylesheet">
                                            <table width="100%" border="0" cellspacing="1" cellpadding="3" class=a2 style="border-collapse: collapse">
                                              <tr>
                                                <td class=a1>
                                                <div align="center">
                                                  <font style="font-size:9pt">
                                                  <b>������Ϣ</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                �����Ը���ĳһ�ض��Ĵʻ���������з��������ӡ��û�����ʱ�估�ض�����̳����в�ѯ��</font></td>
                                              </tr>
                                            </table>
                                          <%end sub

responseend
%>