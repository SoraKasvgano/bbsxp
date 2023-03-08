<%
'BBSXP文件上传类
Class UpFile_Class

Dim Form,File
Private	oUpFileStream	'上传的数据流

'类结束
Private Sub Class_Terminate
	'清除变量及对像
	Form.RemoveAll
	Set Form = Nothing
	File.RemoveAll
	Set File = Nothing
	oUpFileStream.Close
	Set oUpFileStream = Nothing
End Sub

'分析上传的数据
Public Sub GetData ()
	 '定义变量
	Dim RequestBinDate,sSpace,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
	Dim sFormValue,sFileName
	Dim iFindStart,iFindEnd
	Dim iFormStart,iFormEnd,sFormName
	'代码开始
	Set Form = Server.CreateObject ("Scripting.Dictionary")
	Form.CompareMode = 1
	Set File = Server.CreateObject ("Scripting.Dictionary")
	File.CompareMode = 1
	Set tStream = Server.CreateObject ("ADODB.Stream")
	Set oUpFileStream = Server.CreateObject ("ADODB.Stream")
	oUpFileStream.Type = 1
	oUpFileStream.Mode = 3
	oUpFileStream.Open 
	oUpFileStream.Write Request.BinaryRead (Request.TotalBytes)
	oUpFileStream.Position = 0
	RequestBinDate = oUpFileStream.Read 
	iFormEnd = oUpFileStream.Size
	bCrLf = ChrB (13) & ChrB (10)
	'取得每个项目之间的分隔符
	sSpace = MidB (RequestBinDate,1, InStrB (1,RequestBinDate,bCrLf)-1)
	iStart = LenB(sSpace)
	iFormStart = iStart+2
	'分解项目
	Do
		iInfoEnd = InStrB (iFormStart,RequestBinDate,bCrLf & bCrLf)+3
		tStream.Type = 1
		tStream.Mode = 3
		tStream.Open
		oUpFileStream.Position = iFormStart
		oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
		tStream.Position = 0
		tStream.Type = 2
		tStream.CharSet = BBSxpCharset
		sInfo = tStream.ReadText			
		'取得表单项目名称
		iFormStart = InStrB (iInfoEnd,RequestBinDate,sSpace)-1
		iFindStart = InStr (22,sInfo,"name=""",1)+6
		iFindEnd = InStr (iFindStart,sInfo,"""",1)
		sFormName = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
		'如果是文件
		If InStr (45,sInfo,"filename=""",1) > 0 Then
			Set oFileInfo = new FileInfo_Class
			'取得文件属性
			iFindStart = InStr (iFindEnd,sInfo,"filename=""",1)+10
			iFindEnd = InStr (iFindStart,sInfo,""""&vbCrLf,1)
			sFileName = Trim(Mid(sinfo,iFindStart,iFindEnd-iFindStart))
			oFileInfo.FileName = GetFileName(sFileName)
			oFileInfo.FilePath = GetFilePath(sFileName)
			oFileInfo.FileExt = GetFileExt(sFileName)
			iFindStart = InStr (iFindEnd,sInfo,"Content-Type: ",1)+14
			iFindEnd = InStr (iFindStart,sInfo,vbCr)
			oFileInfo.FileMIME = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
			oFileInfo.FileStart = iInfoEnd
			oFileInfo.FileSize = iFormStart -iInfoEnd -2
			oFileInfo.FormName = sFormName
			file.add sFormName,oFileInfo
		else
		'如果是表单项目
			tStream.Close
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			oUpFileStream.Position = iInfoEnd 
			oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet = BBSxpCharset
			sFormValue = tStream.ReadText
			If Form.Exists (sFormName) Then
				Form (sFormName) = Form (sFormName) & ", " & sFormValue
				else
				Form.Add sFormName,sFormValue
			End If
		End If
		tStream.Close
		iFormStart = iFormStart+iStart+2
		'如果到文件尾了就退出
	Loop Until (iFormStart+2) >= iFormEnd 
	RequestBinDate = ""
	Set tStream = Nothing
End Sub

'保存到文件,自动覆盖已存在的同名文件 by yuzi
Public Function SaveToFile(Item,Path)
	Set oFileStream = CreateObject ("ADODB.Stream")
	oFileStream.Type = 1
	oFileStream.Open
	oUpFileStream.Position = File(Item).FileStart
	oUpFileStream.CopyTo oFileStream,File(Item).FileSize
	oFileStream.SaveToFile Path,2
	oFileStream.Close
	Set oFileStream = Nothing
End Function


'取得文件数据
Public Function FileData(Item)
	oUpFileStream.Position = File(Item).FileStart
	FileData = oUpFileStream.Read (File(Item).FileSize)
End Function


'取得文件路径
Public function GetFilePath(FullPath)
  If FullPath <> "" Then
    GetFilePath = Left(FullPath,InStrRev(FullPath, "\"))
    Else
    GetFilePath = ""
  End If
End function

'取得文件名
Public Function GetFileName(FullPath)
  If FullPath <> "" Then
    GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
    Else
    GetFileName = ""
  End If
End function

'取得文件的后缀名
Public Function GetFileExt(FullPath)
  If FullPath <> "" Then
    GetFileExt = LCase(Mid(FullPath,InStrRev(FullPath, ".")+1))
    Else
    GetFileExt = ""
  End If
End function

End Class


'文件属性类
Class FileInfo_Class
Dim FormName,FileName,FilePath,FileSize,FileMIME,FileStart,FileExt
End Class
%>
