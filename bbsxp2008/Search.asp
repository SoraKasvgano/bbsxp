<!-- #include file="Setup.asp" --><%
HtmlTop

ForumID=RequestInt("ForumID")
DateComparer=RequestInt("DateComparer")

if ForumID>0 then
	Set Rs=Execute("Select Moderated from ["&TablePrefix&"Forums] where ForumID="&ForumID&"")
	If Not Rs.eof Then Moderated=Rs("Moderated")
	Rs.Close
	%><!-- #include file="Utility/ForumPermissions.asp" --><%
elseif BestRole=1 then
	PermissionManage=1
end if


if Request("menu")="Result" then
	Keywords=HTMLEncode(Request("Keywords"))

	SortBy=HTMLEncode(Request("SortBy"))
	Item=HTMLEncode(Request("Item"))

	if Keywords="" then error("您没有输入任何查询条件！")
	if Request("VerifyCode")<>Session("VerifyCode") or Session("VerifyCode")="" then error("验证码错误！")

	SQLSearch="Visible=1 and "&Item&" like '%"&Keywords&"%' "
	
	if DateComparer > 0 then SQLSearch=SQLSearch&" and DateDiff("&SqlChar&"d"&SqlChar&",PostTime,"&SqlNowString&") < "&DateComparer&" "

	if ForumID > 0 then SQLSearch=SQLSearch&" and ForumID="&ForumID&" "


	sql="Select * from ["&TablePrefix&"Threads] where "&SQLSearch&" order by ThreadID "&SortBy&""
	Rs.Open sql,Conn,1
		count=Execute("Select count(ThreadID) from ["&TablePrefix&"Threads] where "&SQLSearch&"")(0)    '数据总条数
		if Count=0 then error("对不起，没有找到您要查询的内容")
		PageSetup=SiteConfig("ThreadsPerPage") '设定每页的显示数量
		Rs.Pagesize=PageSetup
		TotalPage=Rs.Pagecount  '总页数
		PageCount = RequestInt("PageIndex")
		if PageCount <1 then PageCount = 1
		if PageCount > TotalPage then PageCount = TotalPage
		if TotalPage>0 then Rs.absolutePage=PageCount '跳转到指定页数
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 搜索结果</div>

<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td width="100%" align="right">共找到了 <b><font color="FF0000"><%=Count%></font></b> 篇相关帖子</td>
	</tr>
</table>

<%if PermissionManage=1 then response.write("<form style='margin:0px' method=POST action=Manage.asp><input type='hidden' name='ForumID' value='"&ForumID&"' />")%>

<table cellspacing="1" cellpadding="5" width="100%" class="CommonListArea">
	<tr class="CommonListTitle">
		<td colspan="5">主题</td>
        <%if PermissionManage=1 then response.write("<td width=50 align=center><input type=checkbox name=chkall onclick='CheckAll(this.form)' value=ON /></td>")%>
	</tr>
	<tr class="CommonListHeader" align="center">
		<td>标题</td>
		<td>作者</td>
		<td>回复</td>
		<td>查看</td>
		<td>最后更新</td>
        <%if PermissionManage=1 then response.write("<td width=50 align=center></td>")%>
	</tr>
<%
i=0
Do While Not RS.EOF and i<pagesetup
i=i+1
ShowThread()
Rs.MoveNext
loop
Rs.Close
%>
</table>

<%if PermissionManage=1 then%>
<table cellspacing=0 cellpadding=5 border=0 width=100%>
	<tr>
		<td align=right>
		管理选项：<select name=menu size=1>
        	<%if BestRole = 1 then%>
			<optgroup label="公告">
				<option value="Top">设为公告</option>
				<option value="UnTop">取消公告</option>
			</optgroup>
            <%end if%>
			<optgroup label="解锁">
				<option value="IsLocked">锁定主题</option>
				<option value="DelIsLocked">解锁主题</option>
			</optgroup>
			<optgroup label="精华">
				<option value="IsGood">精华主题</option>
				<option value="DelIsGood">取消精华</option>
			</optgroup>
			<optgroup label="审核">
				<option value="Visible">审核通过</option>
				<option value="InVisible">审核失败</option>
			</optgroup>
			<optgroup label="置顶">
				<option value="ThreadTop">置顶主题</option>
				<option value="DelTop">取消置顶</option>
			</optgroup>
			<optgroup label="其它">
				<option value="DelThread">删除主题</option>
				<option value="Fix">修复主题</option>
				<option value="MoveNew">拉前主题</option>
				<option value="MoveThread">移动主题</option>
			</optgroup>
		</select>　<input type="submit" value=" 执 行 " onclick="return VerifyRadio('Item');" />
		</td>
</table>
<%
	response.write("</form>")
end if
%>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr>
		<td><%ShowPage()%></td>
	</tr>
</table>

<script language="JavaScript" type="text/javascript">
function VerifyRadio() {
	objYN=false;
	if (window.confirm('您确定执行本次操作?')){
		for (i=0;i<document.getElementsByName("ThreadID").length;i++) {
			if (document.getElementsByName("ThreadID")[i].checked) {objYN= true;}
		}
		if (objYN==false) {alert ('请选择您要操作的主题！');return false;}
	}
	return objYN;
}
</script>
<%
	HtmlBottom
end if
%>
<div class="CommonBreadCrumbArea"><%=ClubTree%> → 搜索帖子</div>
<table cellspacing=1 cellpadding=5 width=100% class=CommonListArea>
	<form method="POST" action="Search.asp?menu=Result" name="form" onsubmit="return doSearch();">
		<tr class=CommonListTitle>
			<td colspan="2">搜索选项</td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%" align="right">关 键 词：</td>
			<td><input size="40" name="Keywords" onkeyup="ValidateTextboxAdd(this, 'btnadd')" /></td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%" align="right">搜索方式：</td>
			<td> 
			<input type="radio" name="searchMethod" value="BBS" id="BBS" checked="checked" /><label for="BBS">站内</label> 
			<input type="radio" name="searchMethod" value="DuoCi" id="DuoCi" /><label for="DuoCi">DuoCi搜索</label> </td>
		</tr>
		<tr class="CommonListCell">
			<td width="40%" align="right">搜索内容：</td>
			<td> 
			<input type="radio" name="Item" value="Topic" id="Topic" checked="checked" /><label for="Topic">主题</label> 
			<input type="radio" name="Item" value="Category" id="Category" /><label for="Category">类别</label> 
			<input type="radio" name="Item" value="PostAuthor" id="PostAuthor" /><label for="PostAuthor">作者</label> 
			<input type="radio" name="Item" value="LastName" id="LastName" /><label for="LastName">最后更新的作者</label> 
		</tr>
		<tr class="CommonListCell">
			<td align="right">日期范围：</td>
			<td>
			<select size="1" name="DateComparer">
				<option value="1" >昨天以来</option>
				<option value="2" >2 天以来</option>
				<option value="7" >1 周以来</option>
				<option value="10" >10 天以来</option>
				<option value="14" >2 周以来</option>
				<option value="30" >1 个月以来</option>
				<option value="45" >45 天以来</option>
				<option value="60" >2 个月以来</option>
				<option value="75" >75 天以来</option>
				<option value="100" >100 天以来</option>
				<option value="365" selected="selected">1 年以来</option>
				<option value="-1">任何时间</option>
				</select>
</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">排序方式：</td>
			<td>
				<select  name="SortBy">
					<option value="">从旧到新</option>
					<option value="Desc" selected="">从新到旧</option>
				</select>
			</td>
		</tr>
		<tr class="CommonListCell">
			<td align="right">搜索论坛：</td>
			<td>
				<select name="ForumID">
				<option value="0">全部论坛</option>
				<%=GroupList(ForumID)%>
				</select>
			</td>
		</tr>
		
	<tr class="CommonListCell">
		<td align="right">验 证 码：</td>
		<td><input type="text" name="VerifyCode" maxlength="4" size="10" onBlur="CheckVerifyCode(this.value)" onKeyUp="if (this.value.length == 4)CheckVerifyCode(this.value)" onfocus="getVerifyCode()" /> <span id="VerifyCodeImgID">点击输入框获取验证码</span> <span id="CheckVerifyCode" style="color:red"></span></td>
	</tr>
		<tr class="CommonListCell">
			<td colspan="2" align="center"><input type="submit" value="开始搜索" /></td>
		</tr>
	</form>
</table>
<script language="javascript" type="text/javascript">
function doSearch() {
	searchMethod = document.form.searchMethod[1].checked;
	Keyword = document.form.Keywords.value;
	if (Keyword == "") {
		alert(请输入搜索关键词);
		document.form.Keywords.focus();
		return false;
	}
	Keyword = "site:<%=Server_Name%> "+Keyword;
	if (searchMethod == true) {
		window.open("http://www.duoci.com/Search/?Charset=<%=BBSxpCharset%>&word="+Keyword);
		return false
	}
	else {
		return true;
	}
}

function VerifyRadio() {
	objYN=false;
	if (window.confirm('您确定执行本次操作?')){
		for (i=0;i<document.getElementsByName("ThreadID").length;i++) {
			if (document.getElementsByName("ThreadID")[i].checked) {objYN= true;}
		}
	
		if (objYN==false) {alert ('请选择您要操作的主题！');return false;}
	}
}
</script>
<%
HtmlBottom
%>