query|||7200|||group|||select top 10 ID,GroupName,GroupInfo,AppUserID,AppUserName,UserNum,Stats,PostNum,TopicNum,TodayNum,YesterdayNum,LimitUser,PassDate,GroupLogo From [Dv_GroupName] Where Stats>0 order by UserNum desc|||<ul>|||<LI>{$GroupName}¥¥Ω®»À{$AppUserName}</LI>
<LI>{$ID}</LI>
<LI>[{$GroupLogo}]</LI>
<LI>{$UserNum}</LI>
<LI>{$Stats}</LI>
<LI>{$TodayNum}</LI>
<LI>{$LimitUser}{$PassDate}</LI>|||</ul>|||10$3$1|||20|||2