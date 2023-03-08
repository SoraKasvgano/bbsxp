<%

dim Jpeg,upfile,FileUP,FormName
dim FileMessage,FileName,FileMIME,FileSize,FileExt,SaveFile,TotalFileSize,Script
TotalFileSize=0
TotalUserPostAttachments=Execute("Select sum(ContentSize) from ["&TablePrefix&"PostAttachments] where UserName='"&CookieUserName&"'")(0)

UploadFile

Sub CheckFileExt()
	if instr("|gif|jpg|jpeg|png|","|"&FileExt&"|") > 0 then
		if split(FileMIME,"/")(0)<>"image" then FileMessage="后缀名与文件类型不符合":Exit Sub
	end if

	if UpClass="Face" then
		if instr("|gif|jpg|jpeg|png|","|"&FileExt&"|") <= 0 then FileMessage="对不起，头像只能上传后缀名为 gif、jpg、jpeg、png 格式的文件":Exit Sub
	else
		if instr("|"&SiteConfig("UpFileTypes")&"|","|"&FileExt&"|") <= 0 then FileMessage="对不起，管理员设定本论坛不允许上传 "&FileExt&" 格式的文件":Exit Sub
	end if
	
	if FileExt="asa" or FileExt="asp" or FileExt="cdx" or FileExt="cer" or FileExt="aspx" then FileMessage="对不起，管理员设定本论坛不允许上传 "&FileExt&" 格式的文件":Exit Sub

	if FileSize < 1 then FileMessage="当前文件为空文件":Exit Sub
	if FileSize > UpMaxFileSize then FileMessage="文件大小不得超过 "&CheckSize(UpMaxFileSize)&"\n当前的文件大小为 "&CheckSize(FileSize)&"":Exit Sub

	if UpClass<>"Face" then
		if TotalUserPostAttachments+FileSize>UpMaxPostAttachmentsSize then FileMessage="您的上传空间已满！":Exit Sub
	end if
End Sub

Sub UploadFile
	if ""&SiteConfig("UpFileOption")&""="" then
		FileMessage="对不起，管理员关闭文件上传功能"
		if UpClass="Face" then Alert(""&FileMessage&"")
	elseif SiteConfig("UpFileOption")="ADODB.Stream" then
		Set upfile=new upfile_class						'建立上传对象
		upfile.GetData ()								'取得上传数据

		i=0
		for each FormName in upfile.file
			if UpClass<>"Face" then UpFileName=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&i&""
			Set EachFile = upfile.file(FormName)
			FileName=EachFile.FileName				'文件名
			FileExt=LCase(EachFile.FileExt)			'小写后缀名
			FileMIME=EachFile.FileMIME				'文件类型
			FileSize=EachFile.FileSize				'文件大小

			CheckFileExt()
			if ""&FileMessage&""<>"" then
				if UpClass="Face" then
					Alert(""&FileMessage&"")
				else
					Exit Sub
				end if
			end if

			SaveFile=""&UpFolder&UpFileName&"."&FileExt&""			'保存文件路径
			if SiteConfig("AttachmentsSaveOption")=1 or UpClass="Face" then upfile.SaveToFile FormName,Server.mappath(""&SaveFile&"")
			Set EachFile=nothing
			if UpClass<>"Face" then AddToDB
			ImagePersits

			i=i+1
		next
		set upfile=nothing
	'''''''''''''''''''''''''''''''''''''''''''''
	elseif SiteConfig("UpFileOption")="SoftArtisans.FileUp" then
		Set FileUP = Server.CreateObject("SoftArtisans.FileUp")
		i=0
		For Each FormName In FileUP.Form
			if UpClass<>"Face" then UpFileName=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&i&""
			If IsObject(FileUP.Form(FormName)) Then
				If Not FileUP.Form(FormName).IsEmpty Then
					FileName = FileUP.Form(FormName).ShortFileName	 		'原文件名
					FileExt=LCase(mid(FileName,InStrRev(FileName, ".")+1))	'小写后缀名
					FileMIME=FileUP.Form(FormName).ContentType				'文件类型
					Filesize = FileUP.Form(FormName).TotalBytes				'文件大小

					CheckFileExt()
					if ""&FileMessage&""<>"" then
						if UpClass="Face" then
							Alert(""&FileMessage&"")
						else
							Exit Sub
						end if
					end if

					SaveFile=""&UpFolder&UpFileName&"."&FileExt&""			'保存文件路径
					if SiteConfig("AttachmentsSaveOption")=1 or UpClass="Face" then FileUP.Form(FormName).SaveAs Server.mappath(""&SaveFile&"")
					
					
					if UpClass<>"Face" then AddToDB
					ImagePersits
				End If
			End If
			i=i+1	
		Next

	'''''''''''''''''''''''''''''''''''''''''''''
	elseif SiteConfig("UpFileOption")="Persits.Upload" then
	Set Upload = Server.CreateObject("Persits.Upload")  
	Upload.Save

		i=0

		For Each File in Upload.Files


			if UpClass<>"Face" then UpFileName=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&i&""

					FileName = File.FileName	 		'原文件名
					FileExt=LCase(mid(FileName,InStrRev(FileName, ".")+1))	'小写后缀名

					FileMIME=File.ContentType				'文件类型
					Filesize =File.Size				'文件大小

					CheckFileExt()
					if ""&FileMessage&""<>"" then
						if UpClass="Face" then
							Alert(""&FileMessage&"")
						else
							Exit Sub
						end if
					end if

					SaveFile=""&UpFolder&UpFileName&"."&FileExt&""			'保存文件路径

					if SiteConfig("AttachmentsSaveOption")=1 or UpClass="Face" then File.Saveas Server.mappath(""&SaveFile&"")
					if UpClass<>"Face" then AddToDB
					ImagePersits

			i=i+1	
		Next
	end if
End Sub



Sub ImagePersits
	if IsObjInstalled("Persits.Jpeg") and SiteConfig("AttachmentsSaveOption")=1 and FileMIME="image/pjpeg" then
		Set Jpeg = Server.CreateObject("Persits.Jpeg")
		Jpeg.Open Server.MapPath(""&SaveFile&"")
	
		if UpClass="Face" then		'上传头像自动缩放到后台设定的指定值
			Jpeg.Width = Jpeg.OriginalWidth
			Jpeg.Height = Jpeg.OriginalHeight
			if Jpeg.OriginalWidth / Jpeg.OriginalHeight >= 1 then 
				if Jpeg.Width>SiteConfig("AvatarWidth") then
					Jpeg.Width = SiteConfig("AvatarWidth")
					Jpeg.Height = int((SiteConfig("AvatarWidth")/Jpeg.OriginalWidth)*Jpeg.OriginalHeight)
				end if
			elseif Jpeg.OriginalWidth / Jpeg.OriginalHeight < 1 then
				if Jpeg.Height>SiteConfig("AvatarHeight") then
					Jpeg.Height = SiteConfig("AvatarHeight")
					Jpeg.Width= int(Jpeg.OriginalWidth*(SiteConfig("AvatarHeight")/Jpeg.OriginalHeight))
				end if
			end if
		else						'设置帖子附件图片水印效果
			if SiteConfig("WatermarkOption")="Persits.Jpeg" then
				JpegPersits
			end if
		end if
		
		Jpeg.Save Server.MapPath(""&SaveFile&"")
		Set Jpeg = nothing
	end if
End Sub


'''''''''''''''''''''写入表 Start'''''''''''''''''''
Sub AddToDB()
	Rs.Open "Select top 1 * from ["&TablePrefix&"PostAttachments]",Conn,1,3
	Rs.addnew 
		Rs("UserName")=CookieUserName
		Rs("FileName")=FileName
		Rs("ContentType")=FileMIME
		Rs("ContentSize")=FileSize
		if SiteConfig("AttachmentsSaveOption")=1 then
			Rs("FilePath")=SaveFile
		else
			if SiteConfig("UpFileOption")="ADODB.Stream" then Rs("FileData")=upfile.FileData(FormName)
			if SiteConfig("UpFileOption")="SoftArtisans.FileUp" then FileUP.SaveAsBlob Rs("FileData")
			if SiteConfig("UpFileOption")="Persits.Upload" then Rs("FileData")=File.Binary
			SaveFile="GetAttachment.asp?AttachmentID="&Rs("UpFileID")&""
		end if
	Rs.update

	Script=Script&"Bbsxp_InsertIntoEdit('"&Rs("UpFileID")&"','"&FileName&"','"&FileMIME&"','"&SaveFile&"');"
	Rs.close
End Sub
'''''''''''''''''''''写入表  End'''''''''''''''''''
%>