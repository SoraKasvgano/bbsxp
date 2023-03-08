<!-- #include file="Setup.asp" -->
<%
AdminTop
if RequestCookies("UserPassword")="" or RequestCookies("UserPassword")<>session("pass") then response.redirect "Admin_Default.asp"
Log("")

ParentID=RequestInt("ParentID")
NodeID=RequestInt("NodeID")
NodeName=HTMLEncode(Request("NodeName"))
NodeUrl=HTMLEncode(Request("NodeUrl"))


If instr(Request("menu"),"Themes")>0 Then
	XmlFilePath=Server.MapPath("Xml/Themes.xml")
ElseIf instr(Request("menu"),"Menu")>0 Then
	XmlFilePath=Server.MapPath("Xml/Menu.xml")
ElseIf instr(Request("menu"),"Emoticon")>0 Then
	width=RequestInt("width")
	height=RequestInt("height")
	rows=RequestInt("rows")
	Columns=RequestInt("Columns")
	XmlFilePath=Server.MapPath("Xml/Emoticons.XML")
End If

Set XMLDOM=Server.CreateObject("Microsoft.XMLDOM")
XMLDOM.load(XmlFilePath)
Set XMLRoot = XMLDOM.documentElement

select case Request("menu")
	case "showThemes"
		showThemes
	case "AddThemes"
		AddThemes
	case "AddThemesok"
		ThemeName=HTMLEncode(Request("ThemeName"))
		if ThemeName="" then Alert("请输入主题名称。")
		if NodeName="" then Alert("请输入主题文件夹名称。")
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if not fso.folderexists(Server.MapPath("Themes/"&NodeName)) then Alert("您输入的主题文件夹名称不存在")
		Set fso = nothing
		
		Set TempNode = XMLDOM.createNode("element","Theme","")
		TempNode.text = ThemeName
		AppendNewAttribute "Name",NodeName
		XMLRoot.appendChild(TempNode)
		
		XMLDOM.save(XmlFilePath)
		Set TempNode = nothing
		showThemes
	case "EditThemes"
		EditThemes
	case "EditThemesok"
		ThemeName=HTMLEncode(Request("ThemeName"))
		if ThemeName="" then Alert("请输入主题名称。")
		if NodeName="" then Alert("请输入主题文件夹名称。")
		Set fso = Server.CreateObject("Scripting.FileSystemObject")
		if not fso.folderexists(Server.MapPath("Themes/"&NodeName)) then Alert("您输入的主题文件夹名称不存在")
		Set fso = nothing
		
		Set EditNode=XMLRoot.childNodes(ParentID)
		EditNode.setAttribute "Name",NodeName
		EditNode.text=ThemeName
		Set EditNode=nothing
		XMLDOM.save(XmlFilePath)
		Response.Redirect("?menu=showThemes")
	case "DelThemes"
		DelNode ParentID,-1
		Response.Redirect("?menu=showThemes")



	case "showMenu"
		ShowMenu
	case "addMenu"
		addMenu
	case "addMenuok"
		If ParentID>=0 Then
			Set TempNode = XMLDOM.createNode("element","Menu","")
			TempNode.text = NodeName
			XMLRoot.childNodes(ParentID).appendChild(TempNode)
		Else
			Set TempNode = XMLDOM.createNode("element","Category","")
			AppendNewAttribute "Name",NodeName
			XMLRoot.appendChild(TempNode)
		End If
		AppendNewAttribute "Url",NodeUrl
		
		XMLDOM.save(XmlFilePath)
		Set TempNode = nothing
		ShowMenu
	case "editMenu"
		editMenu
	case "editMenuok"
		editMenuok
	case "DelMenu"
		DelNode ParentID,NodeID
		Response.Redirect("?menu=showMenu")
		
	case "showEmoticon"
		ShowEmoticon
	case "addEmoticon"
		addEmoticon
	case "addEmoticonok"
		If ParentID>=0 Then		'添加子节点
			Set TempNode = XMLDOM.createNode("element","ICON","")
			TempNode.text = NodeName
			AppendNewAttribute "FileName",NodeUrl
			XMLRoot.childNodes(ParentID).appendChild(TempNode)
		Else					'添加表情组
			Set TempNode = XMLDOM.createNode("element","Emoticon","")
			AppendNewAttribute "CategoryName",NodeName
			AppendNewAttribute "PathName",NodeUrl
			AppendNewAttribute "Width",Width
			AppendNewAttribute "Height",Height
			AppendNewAttribute "TableRow",rows
			AppendNewAttribute "TableCol",Columns
			XMLRoot.appendChild(TempNode)
		End If
		
		XMLDOM.save(XmlFilePath)
		Set TempNode = nothing
		Response.Redirect("?menu=showEmoticon")
	case "EditEmoticon"
		EditEmoticon
	case "EditEmoticonok"
		EditEmoticonok
	case "DelEmoticon"
		DelNode ParentID,NodeID
		Response.Redirect("?menu=showEmoticon")
end select

Function AppendNewAttribute(attributeName,attributeValue)
	Set NewAttribute=XMLDOM.CreateNode("attribute",attributeName,"")
	NewAttribute.Text=attributeValue
	TempNode.SetAttributeNode NewAttribute
End Function

Sub DelNode(ParentID,NodeID)
	Set ParentNode=XMLRoot.childNodes(ParentID)
	If NodeID>=0 Then	'删除子节点
		ParentNode.removeChild ParentNode.childNodes(NodeID)
	Else				'删除你节点
		XMLRoot.removeChild ParentNode
	End If
	XMLDOM.save(XmlFilePath)
	Set ParentNode=nothing
End Sub

Sub ShowThemes
%>
论坛风格管理

<table cellspacing=1 cellpadding=5 width=70% border=0 class=CommonListArea align=center>
	<tr class=CommonListTitle>
		<td align=center>风格</td>
		<td align=right width=150>
			<a href="?menu=AddThemes">添加</a>
		</td>
	</tr>
<%
	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set ParentNode=XMLRoot.childNodes(NodeIndex)
%>
	<tr class="CommonListCell">
		<td><%=ParentNode.text%>（<%=ParentNode.getAttribute("Name")%>）</td>
		<td align=right>
			<a href=?menu=EditThemes&ParentID=<%=NodeIndex%>>编辑</a> | 
			<a href=?menu=DelThemes&ParentID=<%=NodeIndex%> onclick="return window.confirm('确实执行此操作？')">删除</a>
		</td>
	</tr>
<%
	Next
	Set ParentNode=nothing
%>
</table>
<br />
<%
End Sub

Sub AddThemes
%>
<form method="POST" action="?menu=AddThemesok" name=form>
<table cellspacing="1" cellpadding="5" width="80%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center" colspan="2">添加风格</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="30%"><b>主题名称：</b></td>
		<td width="70%"><input name="ThemeName" size=30></td>
	</tr>
    <tr class="CommonListCell">
		<td align="right" width="30%"><b>文件夹名：</b><br />Themes文件夹下的目录名称</td>
		<td width="70%"><input name="NodeName" size=30></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="100%" colspan="2"> 
			<input type="submit" value=" 添 加 "> <input type="reset" value=" 重 填 ">
		</td>
	</tr>
</table>
</form>
<%
end Sub


Sub EditThemes
	Set EditNode=XMLRoot.childNodes(ParentID)
	NodeName=EditNode.getAttribute("Name")
	ThemeName=EditNode.text
	Set EditNode=nothing
%>
<form method="POST" action="?menu=EditThemesok" name=form>
<input type="hidden" name="ParentID" value="<%=ParentID%>" />
<table cellspacing="1" cellpadding="5" width="60%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center" colspan="2">编辑风格</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="20%"><b>主题名称：</b></td>
		<td><input name="ThemeName" value="<%=ThemeName%>" size="50" /></td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="10%"><b>文件夹名：</b><br />Themes文件夹下的目录名称</td>
		<td>
		<input name="NodeName" value="<%=NodeName%>" size="50"></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="100%" colspan="2"> 
		<input type="submit" value=" 编 辑 ">
		<input type="reset" value=" 重 填 "></td>
	</tr>
</table>
</form>
<%
End Sub





Sub ShowMenu
%>
论坛菜单管理<br />
<form method="POST" action="?menu=addMenuok" name=form>
<input type=hidden name=ParentID value=-1>
<input type=hidden name=NodeUrl value="#">
菜单名称：（例如：论坛插件）<input name="NodeName"> <input type="submit" value="添加"></form>
<table cellspacing=1 cellpadding=5 width=70% border=0 class=CommonListArea align=center>
	<%Adminmenu(0)%>
</table>
<br />
<%
End Sub

Sub addMenu
%>
<form method="POST" action="?menu=addMenuok" name=form>
<table cellspacing="1" cellpadding="5" width="60%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center" colspan="4">添加菜单</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="10%">标题：</td>
		<td width="40%"><input name="NodeName"></td>
		<td align="right" width="10%">分类：</td>
		<td width="40%">
		<select name="ParentID">
			<option value="-1">一级菜单</option>
<%
	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set childNode=XMLRoot.childNodes(NodeIndex)
%>
			<option value="<%=NodeIndex%>" <%if ParentID=NodeIndex then%>selected<%end if%>><%=childNode.getAttribute("Name")%></option>
<%
	Next
	Set childNode=nothing
%>
		</select>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="10%">链接：</td>
		<td width="90%" colspan="3"><input name="NodeUrl" size="50"></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="100%" colspan="4"> 
			<input type="submit" value=" 添 加 "> <input type="reset" value=" 重 填 ">
		</td>
	</tr>
</table>
<%
End Sub

Sub editMenuok
	OldParentID=RequestInt("OldParentID")
	If NodeID=Request("ParentID") Then Alert("设置错误")
	
	If ParentID<0 Then
		Set EditNode=XMLRoot.childNodes(NodeID)
		EditNode.setAttribute "Name",NodeName
		EditNode.setAttribute "Url",NodeUrl
		Set EditNode=nothing
	Else
		if OldParentID=ParentID then	'只是编辑
			Set ParentNode=XMLRoot.childNodes(ParentID)
			Set EditNode=ParentNode.childNodes(NodeID)
			EditNode.text=NodeName
			EditNode.setAttribute "Url",NodeUrl
			Set EditNode=nothing
		else							'移动到其它节点，直接增加并删除原有节点。
			Set TempNode = XMLDOM.createNode("element","Menu","")
			TempNode.text = NodeName
			Set NewAttribute=XMLDOM.CreateNode("attribute","Url","")
			NewAttribute.Text=NodeUrl
			TempNode.SetAttributeNode NewAttribute
			XMLRoot.childNodes(ParentID).appendChild(TempNode)
			DelNode OldParentID,NodeID
			Set TempNode = nothing
		end if
	End If
	XMLDOM.save(XmlFilePath)
	Response.Redirect("?menu=showMenu")
End Sub

Sub editMenu
	If ParentID>=0 Then		'编辑子节点
		Set ParentNode=XMLRoot.childNodes(ParentID)
		Set EditNode=ParentNode.childNodes(NodeID)
		NodeName=EditNode.text
	Else
		Set EditNode=XMLRoot.childNodes(NodeID)
		NodeName=EditNode.getAttribute("Name")
	End If
%>
<form method="POST" action="?menu=editMenuok" name=form>
<input type=hidden name=NodeID value=<%=NodeID%> />
<input type="hidden" name="OldParentID" value="<%=ParentID%>" />
<table cellspacing="1" cellpadding="5" width="60%" border="0" class=CommonListArea align="center">
	<tr class=CommonListTitle>
		<td align="center" colspan="4">编辑菜单</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="10%">标题：</td>
		<td width="40%"><input name="NodeName" value="<%=NodeName%>" /></td>
		<td align="right" width="10%">分类：</td>
		<td width="40%">
		<select name="ParentID">
			<option value="-1">一级菜单</option>
<%
	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set childNode=XMLRoot.childNodes(NodeIndex)
%>
			<option value="<%=NodeIndex%>" <%if ParentID=NodeIndex then%>selected<%end if%>><%=childNode.getAttribute("Name")%></option>
<%
	Next
	Set childNode=nothing
%>
		</select>
		</td>
	</tr>
	<tr class="CommonListCell">
		<td align="right" width="10%">链接：</td>
		<td colspan="3">
		<input name="NodeUrl" value="<%=EditNode.getAttribute("Url")%>" size="50"></td>
	</tr>
	<tr class="CommonListCell">
		<td align="center" width="100%" colspan="4"> 
		<input type="submit" value=" 编 辑 ">
		<input type="reset" value=" 重 填 "></td>
	</tr>
</table>
<%
	Set EditNode=nothing
End Sub

Sub AdminMenu(selec)
	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set ParentNode=XMLRoot.childNodes(NodeIndex)
%>
	<tr class=CommonListTitle>
		<td align=center><%=ParentNode.getAttribute("Name")%></td>
		<td align=right width=150>
			<a href="?menu=addMenu&ParentID=<%=NodeIndex%>">添加</a> | 
			<a href="?menu=editMenu&ParentID=-1&NodeID=<%=NodeIndex%>">编辑</a> | 
			<a href="?menu=DelMenu&ParentID=<%=NodeIndex%>&NodeID=-1" onclick="return window.confirm('确实执行此操作？')">删除</a>
		</td>
	</tr>
<%
		For j=0 To ParentNode.childNodes.length-1
			Set childNode=ParentNode.childNodes(j)
%>
	<tr class="CommonListCell">
		<td><%=childNode.text%>（<a href=<%=childNode.getAttribute("Url")%> target=_blank><%=childNode.getAttribute("Url")%></a>）</td>
		<td align=right>
			<a href=?menu=editMenu&ParentID=<%=NodeIndex%>&NodeID=<%=j%>>编辑</a> | 
			<a href=?menu=DelMenu&ParentID=<%=NodeIndex%>&NodeID=<%=j%> onclick="return window.confirm('确实执行此操作？')">删除</a>
		</td>
	</tr>
<%
		Next
		Set childNode=nothing
	Next
	Set ParentNode=nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowEmoticon
%>
<table border="0" width="80%" align="center">
	<tr>
		<td><a href='http://www.biaoqing.com/' target=_blank title='表情网'>获得更多表情</a></td>
		<td align=right><a href="?menu=addEmoticon&ParentID=-1" class="CommonTextButton">添加表情组</a>
</td>
	</tr>
</table>
<br />
<%
	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set ParentNode=XMLRoot.childNodes(NodeIndex)
		TableRow=ParentNode.getAttribute("TableRow")
		TableCol=ParentNode.getAttribute("TableCol")
		Width=ParentNode.getAttribute("Width")
%>
<table cellspacing=1 cellpadding=5 width=80% border=0 class=CommonListArea align=center>
	<tr class=CommonListTitle>
		<td align=center colspan="<%=TableCol%>"><div style="float:left"><%=ParentNode.getAttribute("CategoryName")%></div>
		<div style="float:right">
			<a href="?menu=addEmoticon&ParentID=<%=NodeIndex%>">添加</a> | 
			<a href="?menu=EditEmoticon&ParentID=-1&NodeID=<%=NodeIndex%>">编辑</a> | 
			<a href="?menu=DelEmoticon&ParentID=<%=NodeIndex%>&NodeID=-1" onclick="return window.confirm('确实执行此操作？')">删除</a>
		</div>
		</td>
	</tr>
	
	<tr class="CommonListCell">
<%
		i=0
		For j=0 To ParentNode.childNodes.length-1
			if i>0 and i mod TableCol=0 then response.Write("</tr><tr class=CommonListCell>")
			Set childNode=ParentNode.childNodes(j)
%>
		<td align="center"><a href="<%=ParentNode.getAttribute("PathName")&childNode.getAttribute("FileName")%>" onmouseover="showmenu(event,'<div class=menuitems><a href=?menu=EditEmoticon&ParentID=<%=NodeIndex%>&NodeID=<%=j%>>编辑</a></div><div class=menuitems><a href=?menu=DelEmoticon&ParentID=<%=NodeIndex%>&NodeID=<%=j%>>删除</a></div>')" title="<%=childNode.text%>" target="_blank"><img src="<%=ParentNode.getAttribute("PathName")&childNode.getAttribute("FileName")%>" width="<%=Width%>" border="0" /></a></td>
		

<%
			i=i+1
		Next
		Set childNode=nothing
		if i mod TableCol<>0 then
			for j=1 to TableCol-i mod TableCol
				response.Write("<td></td>")
			next
		end if
%>
	</tr>
</table>
<%
	Next
	Set ParentNode=nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub addEmoticon
%>
<table cellspacing=1 cellpadding=5 width=70% border=0 class=CommonListArea align=center>
<form method="POST" action="?menu=addEmoticonok" name=form>
<input type=hidden name=NodeID value=-1>
<tr class=CommonListTitle>
	<td align=center colspan=2>添加表情(组)</td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>名称：</td>
	<td><input name="NodeName"></td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>类别：</td>
	<td>
	
		<select name="ParentID">
			<%if ParentID<0 then%><option value="-1">------</option><%end if

	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set childNode=XMLRoot.childNodes(NodeIndex)
%>
			<option value="<%=NodeIndex%>" <%if ParentID=NodeIndex then%>selected<%end if%>><%=childNode.getAttribute("CategoryName")%></option>
<%
	Next
	Set childNode=nothing
%>
		</select>
	
	</td>
</tr>
<%if ParentID<0 then%>
<tr class="CommonListCell">
	<td width=25% align=right>目录：</td>
	<td><input name="NodeUrl" size=40></td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>大小：</td>
	<td><input name="Width" size=4>X<input name="Height" size=4> （宽X高）</td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>显示：</td>
	<td><input name="Rows" size=2>行　<input name="Columns" size=2>列</td>
</tr>
<%else%>
<tr class="CommonListCell">
	<td width=25% align=right>文件名：</td>
	<td><input name="NodeUrl" size=40></td>
</tr>
<%end if%>
<tr class="CommonListCell">
	<td align=center colspan=2><input type="submit" value=" 添加 "></td>
</tr>
</form>
</table>
<%
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EditEmoticon
	If ParentID>=0 Then
		Set ParentNode=XMLRoot.childNodes(ParentID)
		Set EditNode=ParentNode.childNodes(NodeID)
		NodeName=EditNode.text
		NodeUrl=EditNode.getAttribute("FileName")
	Else
		Set EditNode=XMLRoot.childNodes(NodeID)
		NodeName=EditNode.getAttribute("CategoryName")
		NodeUrl=EditNode.getAttribute("PathName")
		Width=EditNode.getAttribute("Width")
		Height=EditNode.getAttribute("Height")
		Rows=EditNode.getAttribute("TableRow")
		Columns=EditNode.getAttribute("TableCol")
	End If
%>
<table cellspacing=1 cellpadding=5 width=70% border=0 class=CommonListArea align=center>
<form method="POST" action="?menu=EditEmoticonok" name=form>
<input type=hidden name=NodeID value="<%=NodeID%>">
<input type="hidden" name="OldParentID" value="<%=ParentID%>" />
<tr class=CommonListTitle>
	<td align=center colspan=2>编辑表情(组)</td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>名称：</td>
	<td><input name="NodeName" value="<%=NodeName%>"></td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>类别：</td>
	<td>
	
		<select name="ParentID">
			<%if ParentID<0 then%><option value="-1">------</option><%end if

	For NodeIndex=0 To XMLRoot.childNodes.length-1
		Set childNode=XMLRoot.childNodes(NodeIndex)
%>
			<option value="<%=NodeIndex%>" <%if ParentID=NodeIndex then%>selected<%end if%>><%=childNode.getAttribute("CategoryName")%></option>
<%
	Next
	Set childNode=nothing
%>
		</select>
	
	</td>
</tr>
<%if ParentID<0 then%>
<tr class="CommonListCell">
	<td width=25% align=right>目录：</td>
	<td><input name="NodeUrl" size=40 value="<%=NodeUrl%>"></td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>大小：</td>
	<td><input name="Width" size=5 value="<%=Width%>">X<input name="Height" size=5 value="<%=Height%>"> （宽X高）</td>
</tr>
<tr class="CommonListCell">
	<td width=25% align=right>显示：</td>
	<td><input name="Rows" size=5 value="<%=Rows%>">行　<input name="Columns" size=5 value="<%=Columns%>">列</td>
</tr>
<%else%>
<tr class="CommonListCell">
	<td width=25% align=right>文件名：</td>
	<td><input name="NodeUrl" size=40 value="<%=NodeUrl%>"></td>
</tr>
<%end if%>
<tr class="CommonListCell">
	<td align=center colspan=2><input type="submit" value=" 编辑 "></td>
</tr>
</form>
</table>
<%
End Sub
'''''''''''''''''''''''''''''''''''''''''
Sub EditEmoticonok
	OldParentID=RequestInt("OldParentID")
	If NodeID=Request("ParentID") Then Alert("设置错误")
	
	If ParentID<0 Then
		Set EditNode=XMLRoot.childNodes(NodeID)
		EditNode.setAttribute "CategoryName",NodeName
		EditNode.setAttribute "PathName",NodeUrl
		EditNode.setAttribute "Width",width
		EditNode.setAttribute "Height",height
		EditNode.setAttribute "TableRow",rows
		EditNode.setAttribute "TableCol",Columns
		Set EditNode=nothing
	ElseIf ParentID>=0 Then
		if OldParentID=ParentID then	'只是编辑
			Set ParentNode=XMLRoot.childNodes(ParentID)
			Set EditNode=ParentNode.childNodes(NodeID)
			EditNode.text=NodeName
			EditNode.setAttribute "FileName",NodeUrl
			Set EditNode=nothing
		else							'移动到其它节点，直接增加并删除原有节点。
			Set TempNode = XMLDOM.createNode("element","ICON","")
			TempNode.text = NodeName
			Set NewAttribute=XMLDOM.CreateNode("attribute","FileName","")
			NewAttribute.Text=NodeUrl
			TempNode.SetAttributeNode NewAttribute
			XMLRoot.childNodes(ParentID).appendChild(TempNode)
			DelNode OldParentID,NodeID
			Set TempNode = nothing
		end if
	End If
	XMLDOM.save(XmlFilePath)
	Response.Redirect("?menu=showEmoticon")
End Sub


Set XMLRoot=nothing
Set XMLDOM=nothing
AdminFooter
%>