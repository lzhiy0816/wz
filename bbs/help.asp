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
<title>帮助中心</title>
<table width=97% align="center" border="0">
  <tr>
    <td vAlign="top" width="30%"><a href="http://free.glzc.net/lzhiy0816/"><img src="images/lzylogo.gif" border="0" alt="回首页"></a></td>
    <td vAlign="center" align="top">　<img src="images/closedfold.gif" border="0" />　<a href=main.asp><%=clubname%></a><br />
    　<img src="images/bar.gif" border="0" /><img src="images/openfold.gif" border="0" />　帮助中心</td>
    </tr>
  </table>


    <center>
    <br>

<SCRIPT>valigntop()</SCRIPT>
    <table  cellspacing="1" cellPadding="4" width="97%" border="0" class=a2>
      <tr>
        <td vAlign="top" width="193" class=a3><br>
        帮助菜单<br>
        <a target="I2" href="help.asp?menu=guizhe">社区发贴规则</a><br />
        <a target="I2" href="help.asp?menu=level">等级介绍</a><br />
        <a target="I2" href="help.asp?menu=zhuce">我是否必须注册</font></a><br />
        <a target="I2" href="help.asp?menu=daima">是否可以加入HTML代码 
        </font></a><br />
        <a target="I2" href="help.asp?menu=YBB">YBB代码 </font></a>
        <br />
        <a target="I2" href="help.asp?menu=bz">什么是版主</font></a><br />
        <a target="I2" href="help.asp?menu=cook">使用COOKIE吗</font></a><br />
        <a target="I2" href="help.asp?menu=edit">怎样修改我的注册信息</font></a><br />
        <a target="I2" href="help.asp?menu=gb">短讯息</font></a><br />
        <a target="I2" href="help.asp?menu=cz">能进行查找吗 </font></a>
        <br />
        </td>
        <td vAlign="top" width="541" class=a4>
        <p>
        <iframe name="I2" width="100%" height="250" marginwidth="1" marginheight="1" border="0" frameborder="0">
        浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe></p>
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
            <p align="center"><a><b>　 论坛等级简介</b></a></p>
            </td>
          </tr>
          <tr bgcolor="A09CFF">
            <td bgcolor="FFFFFF" width="106" align="center">等级名称</td>
            <td bgcolor="FFFFFF" width="237">
            <p align="center">经验值</p>
            </td>
            <td bgcolor="FFFFFF" width="155">星级</td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF">
            
           
            
            <font face="宋体"> <script>document.write(""+level(0,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">0 - 100</td>
            <td width="155" bgcolor="FFFFFF">
            <p><font face="宋体">星级为 </font>
<img border="0" src="images/level1.gif" />            </p>
            </td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(101,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">101 - 500</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level2.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(501,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">501 - 1300</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level3.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(1301,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">1301 - 2600</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level4.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(2601,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">2601 - 4500</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level5.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(4501,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">4501 - 7000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level6.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(7001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">7001 - 11000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level7.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(11001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">11001 - 19000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level8.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(19001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">19001 - 30000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level9.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(30001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">30001 - 45000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level10.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(45001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">45001 - 65000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level11.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(65001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">65001 - 90000</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level12.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体"><script>document.write(""+level(90001,1,"","")+levelname+"");</script>：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">90001 - 无限</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level13.gif" width="110" height="11" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体">特邀嘉宾：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level14.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体">版 主：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level15.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体">管 理 员：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
            <img border="0" src="images/level16.gif" /></td>
          </tr>
          <tr>
            <td width="106" align="center" bgcolor="FFFFFF"><font face="宋体">社区区长：</font></td>
            <td width="237" align="center" bgcolor="FFFFFF">--</td>
            <td width="155" bgcolor="FFFFFF"><font face="宋体">星级为 </font>
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
            <font color="FFFFFF">　&nbsp; 社区发贴规则</font></b></a></font></p>
            </td>
          </tr>
          <tr bgcolor="A09CFF">
            <td bgcolor="FFFFFF" width="498" height="1">
            <p align="left">1、经验值：可以从负无穷大到正无穷大，不限。<br />
            　　&nbsp; 金币值：可以从负无穷大到正无穷大，不限。<br />
            　　&nbsp; 体力值：<font color="FF0000">0--100</font></p>
            </td>
            </tr>
            <tr bgcolor="A09CFF">
              <td bgcolor="FFFFFF" width="498" height="67">2、第一次登录到社区初始经验值为0，以后累加<br />
              　　&nbsp; 第一次登录到社区初始金币值为0，以后累加<br />
              　　&nbsp; 第一次登录到社区体力值设为为100，以后累计</td>
              </tr>
              <tr bgcolor="A09CFF">
                <td bgcolor="FFFFFF" width="498" height="1">3、写文章：<br />
                　　&nbsp; 经验值: <font color="FF0000">+5</font><br />
                　　&nbsp; 金币值: <font color="FF0000">+5</font><br />
                　　&nbsp; 体力值: <font color="FF0000">-5</font> （特邀嘉宾、版主、管理员和社区区长不减）</td>
                </tr>
                <tr bgcolor="A09CFF">
                  <td bgcolor="FFFFFF" width="498" height="1">4、回文章：<br />
                  　　&nbsp; 回复者经验值: <font color="FF0000">+2</font><br />
                  　　&nbsp; 回复者金币值: <font color="FF0000">+2</font><br />
                  　　&nbsp; 回复者体力值: <font color="FF0000">-2</font> （特邀嘉宾、版主、管理员和社区区长不减）</td>
                  </tr>
                      <tr bgcolor="A09CFF">
                        <td bgcolor="FFFFFF" width="498" height="1">5、有效停留时间（需要登录社区）：<br />
                      　　&nbsp; 每有效时间<font color="FF0000"> 10 </font>分钟： 经验值:
                      <font color="FF0000">+1</font> 体力值：<font color="FF0000">+10</font></td>
                      </tr>
                      <tr bgcolor="A09CFF">
                        <td bgcolor="FFFFFF" width="498" height="1">6、经验值 &lt;
                        <font color="FF0000">1</font> 或者体力值 &lt;
                        <font color="FF0000">5</font> 者，都不能发表文章。</td>
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
                              <a name="register1">注册</a></b></font></div>
                            </td>
                          </tr>
                          <tr>
                            <td bgcolor="FFFFFF"><font style="font-size:9pt">是的您必须注册才能在发表您的文章！ 您只有正式注册，您才能拥有显示自己个人资料等功能！</font></td>
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
                              <font style="font-size:9pt">使用</font><font style="font-size:9pt" face="Verdana, Arial">HTML</font><font style="font-size:9pt">代码或论坛专用代码</font></b></a></div>
                            </td>
                          </tr>
                          <tr>
                            <td bgcolor="FFFFFF"><font style="font-size:9pt">您不可以在您的消息中使用</font><font style="font-size:9pt" face="Verdana, Arial">HTML</font><font style="font-size:9pt">代码但可以使用论坛专用代码<font face="Verdana, Arial">YBB</font>代码</font><font style="font-size:9pt" face="Verdana, Arial">.&nbsp;
                            </font><font style="font-size:9pt">论坛专用代码类似与</font><font style="font-size:9pt" face="Verdana, Arial">HTML</font><font style="font-size:9pt">代码</font><font style="font-size:9pt" face="Verdana, Arial">,
                            </font><font style="font-size:9pt">但只提供几个基础功能和一些专用功能</font><font style="font-size:9pt" face="Verdana, Arial">,
                            </font><font style="font-size:9pt">如</font><font style="font-size:9pt" face="Verdana, Arial">hyperlinking, 
                            image </font><font style="font-size:9pt">显示</font><font style="font-size:9pt" face="Verdana, Arial">,
                            </font><font style="font-size:9pt">粗体和斜体字</font><font style="font-size:9pt" face="Verdana, Arial">.
                            </font><font style="font-size:9pt">
                            <a target="I2" href="help.asp?menu=YBB">这里</a>是完整的论坛专用代码列表</font><font style="font-size:9pt" face="Verdana, Arial">.
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
                            <b>什么是YBB(Yuzi 
                            Bulletin Board)代码？</font>
                            </font></b></td>
                          </tr>
                          <tr>
                            <td height="58" bgcolor="F7F7F7">
                            <p><font face="Verdana, Arial">YBB</font><font style="font-size:9pt">代码是</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">的一个变种。您也许已经对它很熟悉了。一般情况下，如果允许您用</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">，也就可以使用</font><font face="Verdana, Arial" style="font-size:9pt">YBB</font><font style="font-size:9pt">代码。即使您的讨论区不能让您使用</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">，您也可以使用</font><font face="Verdana, Arial" style="font-size:9pt">YBB</font><font style="font-size:9pt">代码。 
                            由于要求使用的编码很少，即使可以使用</font><font face="Verdana, Arial" style="font-size:9pt">HTML</font><font style="font-size:9pt">，您可能也想使用</font><font face="Verdana, Arial" style="font-size:9pt">YBB</font><font style="font-size:9pt">代码，因为代码错误的可能性大大减小了。</font></p>
                            </td>
                          </tr>
                          <tr>
                            <td class=a1 height="16">
                            <b>

                            URL超级链接</b></td>
                          </tr>
                          <tr>
                            <td height="93" bgcolor="FFFFFF">在您的信息里加入超级链接，只要按下列方式套入就可以了</font><font face="Verdana, Arial">(YBB</font>代码是<font color="ff0000">红字</font></font><font face="Verdana, Arial">).
                            </font><br />
                            </font></font><font face="Verdana, Arial">
                            <br />
                            </font><font face="Verdana, Arial">
                            <font color="ff0000">[url]</font>http://www.yuzi.net<font color="ff0000">[/url]</font>
                            <br />
                            </font></font>或者 <br />
                            <font face="Verdana, Arial"><font color="ff0000">[url=</font>http://www.yuzi.net]Yuzi</font>工作室<font color="ff0000">[/url]</font> <br />
                            </font></font><font><br />
                            按上例套入，</font><font face="Verdana, Arial">YBB</font>代码会自动对</font><font face="Verdana, Arial">URL</font>产生链接，并保证当用户点击新的窗口时这个链接是打开着的。注意</font><font face="Verdana, Arial">URL</font>的</font><font face="Verdana, Arial">&quot;http://&quot;</font>这一部分是随意的。</font></font></td>
                            </tr>
                            <tr>
                              <td class=a1 height="13">
                              <b>电子邮件链接</font></b></td>
                            </tr>
                            <tr>
                              <td height="39" bgcolor="FFFFFF">在您的信息里加入电子邮件的超级链接，只要按照下例套入就可以了(YBB代码是红字)<br />
                              </font></font>
                              <font color="ff0000"><br />
                              <font color="FF0000">[email]</font></font><font color="000000"><a href="mailto:huangzhiyu@yuzi.net">huangzhiyu@yuzi.net</a></font><font color="ff0000">[/email]</font><br />
                              </font></font>
                              <font color="ff0000"><br />
                              </font>按上例套入，YBB代码会对电子邮件自动产生链接。 </font></td>
                              </tr>
                              <tr>
                                <td class=a1 height="13">
                                <b>粗体与斜体</font></b></td>
                              </tr>
                              <tr>
                                <td height="74" bgcolor="FFFFFF"><font>您可以使用 
                                </font><font face="Verdana, Arial">[b] 
                                [/b] </font>或 </font>
                                <font face="Verdana, Arial">[i] [/i]
                                </font>这些标志以达到在帖子中使用粗体或斜体的效果<font face="Verdana, Arial">.</font>
                                </font></font><br />
                                </font></font><br />
                                您好<font face="Verdana, Arial">,
                                <font color="ff0000">[b]</font></font></font><b>管理员</font></font></b><font color="FF0000">[/b]</font><font face="Verdana, Arial"><br />
                                </font>
                                您好</span></font><font face="Verdana, Arial">,
                                <font color="ff0000">[i]</font></font><i>管理员</font></i><font color="ff0000">[/i]</font><font face="Verdana, Arial"><br />
                                </font>您好,<font color="ff0000">[u]</font><u>管理员</u><font color="ff0000">[/u]</font></font></span>
                                </font>
                                <font face="Verdana, Arial"><br />
                                </font>您好,<font color="ff0000">[strike]</font><strike>管理员</strike></font><font color="ff0000">[/strike]</font></span></td>
                                </tr>
                                <tr>
                                  <td class=a1 height="13">
                                  <b>移动文字
                                  </font></b></td>
                                </tr>
                                <tr>
                                  <td height="42" bgcolor="FFFFFF">在您的信息里加入移动文字，只要按下例套入就可以了</font><font face="Verdana, Arial">(YBB</font>代码是<font color="ff0000">红字</font></font><font face="Verdana, Arial">).</font>
                                  </font><font color="ff0000">
                                  <br />
                                  </font><br />
                                  </font></font>
                                  <font color="ff0000">[marquee]</font>移动文字</font><font color="ff0000">[/marquee]</font><br />
                                  </font></font>
                                  <marquee>移动文字</marquee></font></td>
                                  </tr>
                                  <tr>
                                    <td class=a1 height="13">
                                    <b>引用其他信息
                                    </font></b></td>
                                  </tr>
                                  <tr>
                                    <td height="58" bgcolor="FFFFFF">引用一些人的帖子，只要剪切和拷贝然后按下例套入就可以了</font><font face="Verdana, Arial">(YBB</font>代码是 
                                    </font>
                                    <font color="ff0000">红字</font><font face="Verdana, Arial" color="ff0000">)</font><font face="Verdana, Arial">).</font> <br />
                                    </font><font color="ff0000"><br />
                                    [QUOTE]</font>别问您的国家能为您作什么......<br />
                                    问您能为您的国家作什么？</font><font color="ff0000">[/QUOTE]</font>
                                    <br />
                                    <br />
                                    在上例中，</font><font face="Verdana, Arial">YBB</font></font>代码会自动加重您引用的文字。 
                                    </font></td>
                                    </tr>
                                      <tr>
                                        <td height="16" class=a1>
                                        <b>彩色文字</font></b></td>
                                      </tr>
                                      <tr>
                                        <td height="169" bgcolor="FFFFFF">
                                        <p>
                                        <font color="ff0000" face="Verdana, Arial">[color=red]</font>彩色文字<font color="ff0000">[/color]</font><br />
                                        以上的&quot;red&quot;可以是：</font></p>
                                        <div align="center">
                                          <center>
                                          <table border="0" width="49%">
                                            <tr>
                                              <td height=25 width="31%" align="center" bgcolor="F7F7F7">
                                              Black </td>
                                              <td height=25 width="35%" align="center" bgcolor="F7F7F7">
                                              #000000</td>
                                              <td height=25 width="34%" align="center" bgcolor="F7F7F7">
                                              黑色 </td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Silver </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #C0C0C0</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              银色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Gray </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #808080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              灰色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Pink </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FFC8CB</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              粉红</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Maroon </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #800000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              茶色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Red </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FF0000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              红色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Purple </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #800080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              紫色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Fuchsia </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FF00FF</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              紫红</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Green</td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #008000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              绿色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Lime </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #00FF00</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              亮绿</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Olive </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #808000</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              橄榄</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Yellow </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FFFF00</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              黄色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Navy </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #000080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              深蓝</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Blue </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #0000FF</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              蓝色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Teal </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #008080</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              青色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Aqua </td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #00FFFF</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              浅绿</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Orange</td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #FFA500</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              橙色</td>
                                            </tr>
                                            <tr>
                                              <td width="31%" align="center" bgcolor="F7F7F7">
                                              Brown</td>
                                              <td width="35%" align="center" bgcolor="F7F7F7">
                                              #A52A2A</td>
                                              <td width="34%" align="center" bgcolor="F7F7F7">
                                              褐色</td>
                                            </tr>
                                          </table>
                                          </center>
                                        </div>
                                        </td>
                                        </tr>
                                          <tr>
                                            <td height="14" class=a1><b>特别注意</b></td>
                                          </tr>
                                          <tr>
                                            <td height="76" bgcolor="FFFFFF">
                                            <p>您不可以同时使用</font><font face="Verdana, Arial">HTML</font>和</font><font face="Verdana, Arial">YBB</font>代码的同一种功能。并且注意</font><font face="Verdana, Arial">YBB</font>代码不对大小写敏感。所以您可以用</font><font color="ff0000">[URL]</font>
                                            或 </font>
                                            <font color="ff0000">[url]</font>
                                            </font>
                                            <font color="800000">
                                            <br />
                                            不正确的</font><font color="800000" face="Verdana, Arial">YBB</font><font color="800000">代码使用：</font><font face="Verdana, Arial"><font color="ff0000"><br />
                                            [url]</font> www.yuzi.net
                                            <font color="ff0000">[/url]</font>
                                            </font>不要在括号和您输入的文字之间有空格</font><font face="Verdana, Arial"><font color="ff0000"><br />
                                            [email]</font>huangzhiyu@yuzi.net<font color="ff0000">[email]</font>
                                            </font>在结束时，不要忘了在括号内加入斜杠</font><font color="ff0000">[/email]</font>
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
                                                  <b>版主</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                版主就是讨论区的管理员。他可以删除讨论区里的任何帖子。您如果对某讨论区有任何疑问，请直接与其版主联系。</font>
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
                                                本论坛使用了cookie技术，可以存储您的用户名和口令。cookie信息存储在您的浏览器里。cookie不会对您跟踪或产生其它作用，它们使您能更加方便地使用本论坛系统。如果您的浏览器不支持cookie或着您将浏览器的cookie功能关闭，这些省时的功能将无法工作。</font>
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
                                                  <b>更改您的信息</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                您可以通过页面上方的</font><font style="font-size:9pt" face="Verdana, Arial">&quot;控制面板&quot;,
                                                </font>
                                                <font style="font-size:9pt">很方便的更改您的注册信息</font><font style="font-size:9pt" face="Verdana, Arial">.
                                                </font>
                                                <font style="font-size:9pt">输入您的用户名及密码</font><font style="font-size:9pt" face="Verdana, Arial">,
                                                </font>
                                                <font style="font-size:9pt">您将可以查看和修改您先前所有注册的信息</font><font style="font-size:9pt" face="Verdana, Arial">,
                                                </font>
                                                <font style="font-size:9pt">当然除了您的用户名</font><font style="font-size:9pt" face="Verdana, Arial">.
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
                                                  <b>短讯息</b>
                                                  </font>
                                                </div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td height="29" bgcolor="FFFFFF">
                                                <font style="font-size:9pt">当您不在的时候，别人可以发短讯息给您！其他用户是不能看到您的讯息，只有您能够管理自己的讯息。但每个人都可以给您留言！</font></td>
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
                                                  <b>查找消息</b></font></div>
                                                </td>
                                              </tr>
                                              <tr>
                                                <td bgcolor="FFFFFF"><font style="font-size:9pt">
                                                您可以根据某一特定的词或句子在所有发布的帖子、用户名、时间及特定的论坛里进行查询。</font></td>
                                              </tr>
                                            </table>
                                          <%end sub

responseend
%>