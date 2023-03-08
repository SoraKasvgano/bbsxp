<!-- #include file="Setup.asp" -->
<%
if CookieUserName=empty then AlertForModal("您还未登录论坛")
CommentFor=HTMLEncode(Request("CommentFor"))
if CommentFor="" then AlertForModal("没有评论对象")
if Lcase(CommentFor)=Lcase(CookieUserName) then AlertForModal("不能对自己作出评论")


if CookieTotalPosts < SiteConfig("MinReputationPost") then AlertForModal("发帖数少于 "&SiteConfig("MinReputationPost")&" ，无法对他人进行评论")
if CookieReputation < SiteConfig("MinReputationCount") then AlertForModal("声望少于 "&SiteConfig("MinReputationCount")&"，无法对他人进行评论")

if not Execute("select * from ["&TablePrefix&"Reputation] where CommentFor='"&CommentFor&"' and IPAddress='"&REMOTE_ADDR&"' and DateDiff("&SqlChar&"d"&SqlChar&",DateCreated,"&SqlNowString&")=0").eof then AlertForModal("此IP今天已对 "&CommentFor&" 评价过！")


if BestRole<>1 then

	ReputationToday=Execute("Select count(ReputationID) from ["&TablePrefix&"Reputation] where DateDiff("&SqlChar&"d"&SqlChar&",DateCreated,"&SqlNowString&")=0 and CommentBy='"&CookieUserName&"'")(0)
	if ReputationToday>SiteConfig("MaxReputationPerDay") then AlertForModal("每个用户每天对他人的评价不能超过 "&SiteConfig("MaxReputationPerDay")&" 条")

	if SiteConfig("ReputationRepeat")>0 then
	CommentByGetRows=FetchEmploymentStatusList("Select top "&SiteConfig("ReputationRepeat")&" CommentFor from ["&TablePrefix&"Reputation] where CommentBy='"&CookieUserName&"' order by DateCreated DESC")
	if IsArray(CommentByGetRows) then
		For i=0 To Ubound(CommentByGetRows,2)
			if CommentByGetRows(0,i)=CommentFor then  AlertForModal("再次对 "&CommentFor&" 进行声望评价之前，您必须对其他"&SiteConfig("ReputationRepeat")&"个用户进行声望评价！")
		Next
	End if
	CommentByGetRows=null
	end if
	
end if


if Request_Method = "POST" then
	Reputation=RequestInt("Reputation")
	Comment=HTMLEncode(Request("Comment"))
	if Reputation>1 or Reputation<-1 then AlertForModal("非法操作")
	if len(Comment)>500 then AlertForModal("评论内容不能超过500个字符！")
	
	
	if BestRole=1 then Reputation=SiteConfig("AdminReputationPower")*Reputation
	Rs.open "Select * from ["&TablePrefix&"Reputation]",Conn,1,3
	Rs.addnew
		Rs("Reputation")=Reputation
		Rs("Comment")=Comment
		Rs("CommentFor")=CommentFor
		Rs("CommentBy")=CookieUserName
		Rs("IPAddress")=REMOTE_ADDR
	Rs.update
	Rs.close
	
	Execute("Update ["&TablePrefix&"Users] set Reputation=Reputation+"&Reputation&" where UserName='"&CommentFor&"'")
	
	AddApplication "Message_"&CommentFor,"【系统讯息】<a target=_blank href=Profile.asp?UserName="&CommentFor&">"&CookieUserName&" 对您进行了声望评价</a>"

	
%>
<script language="JavaScript" type="text/javascript">
	parent.BBSXP_Modal.Close();
</script>
<%
else
	Response.clear
%>
<title>对 <%=CommentFor%> 进行评论</title>
<style type="text/css">body,table{FONT-SIZE:9pt;}</style>
<script language="JavaScript" type="text/javascript">
	function CheckComment(){
		if (document.form.elements["Reputation"][0].checked != true && document.form.elements["Reputation"][1].checked != true && document.form.elements["Reputation"][2].checked != true){
			alert("请选择总体评价！");
			return false;
		}
		if (document.form.Comment.value == ""){
			alert("请输入您的评论！");
			return false;
		}
		
		if (document.form.Comment.value.length > 500){
			alert("评论内容不能超过500个字符！");
			return false;
		}			
		return true;
	}
</script>
<table width="100%" border=0 align="center">
<form name=form method=Post action=? onsubmit="return CheckComment()">
<input name="CommentFor" type="hidden" value="<%=CommentFor%>" />
  <tr>
    <td height=30>评 价：</td>
    <td>
		<input name="Reputation" type="radio" value="1" id="IsGood" /><label for="IsGood"><img src="images/Reputation_Excellent.gif" align="absmiddle" title="好评" />好评</label>　
		<input name="Reputation" type="radio" value="0" id="IsMid" /><label for="IsMid"><img src="images/Reputation_Average.gif" align="absmiddle" title="中评" />中评</label>　
		<input name="Reputation" type="radio" value="-1" id="IsBad" /><label for="IsBad"><img src="images/Reputation_Poor.gif" align="absmiddle" title="差评" />差评</label>	</td>
  </tr>
  <tr>
    <td valign=top>评 论：</td>
    <td><textarea name="Comment" cols="54" rows="5"></textarea></td>
  </tr>
  <tr>
    <td colspan="2" align="center">
      <input type="submit" name="Submit" value=" 确 定 " />　
      <input type="button" onclick="javascript:parent.BBSXP_Modal.Close()" value=" 取 消 ">
    </td>
    </tr>
</form>
</table>
<%
end if

%>