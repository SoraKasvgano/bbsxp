<%

dim Jpeg,upfile,FileUP,FormName
dim FileMessage,FileName,FileMIME,FileSize,FileExt,SaveFile,TotalFileSize,Script
TotalFileSize=0
TotalUserPostAttachments=Execute("Select sum(ContentSize) from ["&TablePrefix&"PostAttachments] where UserName='"&SqlString(CookieUserName)&"'")(0)

UploadFile

Sub CheckFileExt()
	if instr("|gif|jpg|jpeg|png|","|"&FileExt&"|") > 0 then
		if split(FileMIME,"/")(0)<>"image" then FileMessage="��׺�����ļ����Ͳ�����":Exit Sub
	end if

	if UpClass="Face" then
		if instr("|gif|jpg|jpeg|png|","|"&FileExt&"|") <= 0 then FileMessage="�Բ���ͷ��ֻ���ϴ���׺��Ϊ gif��jpg��jpeg��png ��ʽ���ļ�":Exit Sub
	else
		if instr("|"&SiteConfig("UpFileTypes")&"|","|"&FileExt&"|") <= 0 then FileMessage="�Բ��𣬹���Ա�趨����̳�������ϴ� "&FileExt&" ��ʽ���ļ�":Exit Sub
	end if
	
	if FileExt="asa" or FileExt="asp" or FileExt="cdx" or FileExt="cer" or FileExt="aspx" then FileMessage="�Բ��𣬹���Ա�趨����̳�������ϴ� "&FileExt&" ��ʽ���ļ�":Exit Sub

	if FileSize < 1 then FileMessage="��ǰ�ļ�Ϊ���ļ�":Exit Sub
	if FileSize > UpMaxFileSize then FileMessage="�ļ���С���ó��� "&CheckSize(UpMaxFileSize)&"\n��ǰ���ļ���СΪ "&CheckSize(FileSize)&"":Exit Sub

	if UpClass<>"Face" then
		if TotalUserPostAttachments+FileSize>UpMaxPostAttachmentsSize then FileMessage="�����ϴ��ռ�������":Exit Sub
	end if
End Sub

Sub UploadFile
	if ""&SiteConfig("UpFileOption")&""="" then
		FileMessage="�Բ��𣬹���Ա�ر��ļ��ϴ�����"
		if UpClass="Face" then Alert(""&FileMessage&"")
	elseif SiteConfig("UpFileOption")="ADODB.Stream" then
		Set upfile=new upfile_class						'�����ϴ�����
		upfile.GetData ()								'ȡ���ϴ�����

		i=0
		for each FormName in upfile.file
			if UpClass<>"Face" then UpFileName=""&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&i&""
			Set EachFile = upfile.file(FormName)
			FileName=EachFile.FileName				'�ļ���
			FileExt=LCase(EachFile.FileExt)			'Сд��׺��
			FileMIME=EachFile.FileMIME				'�ļ�����
			FileSize=EachFile.FileSize				'�ļ���С

			CheckFileExt()
			if ""&FileMessage&""<>"" then
				if UpClass="Face" then
					Alert(""&FileMessage&"")
				else
					Exit Sub
				end if
			end if

			SaveFile=""&UpFolder&UpFileName&"."&FileExt&""			'�����ļ�·��
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
					FileName = SafeFileName(FileUP.Form(FormName).ShortFileName)	 		'ԭ�ļ���
					FileExt=LCase(mid(FileName,InStrRev(FileName, ".")+1))	'Сд��׺��
					FileMIME=FileUP.Form(FormName).ContentType				'�ļ�����
					Filesize = FileUP.Form(FormName).TotalBytes				'�ļ���С

					CheckFileExt()
					if ""&FileMessage&""<>"" then
						if UpClass="Face" then
							Alert(""&FileMessage&"")
						else
							Exit Sub
						end if
					end if

					SaveFile=""&UpFolder&UpFileName&"."&FileExt&""			'�����ļ�·��
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

					FileName = SafeFileName(File.FileName)	 		'ԭ�ļ���
					FileExt=LCase(mid(FileName,InStrRev(FileName, ".")+1))	'Сд��׺��

					FileMIME=File.ContentType				'�ļ�����
					Filesize =File.Size				'�ļ���С

					CheckFileExt()
					if ""&FileMessage&""<>"" then
						if UpClass="Face" then
							Alert(""&FileMessage&"")
						else
							Exit Sub
						end if
					end if

					SaveFile=""&UpFolder&UpFileName&"."&FileExt&""			'�����ļ�·��

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
	
		if UpClass="Face" then		'�ϴ�ͷ���Զ����ŵ���̨�趨��ָ��ֵ
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
		else						'�������Ӹ���ͼƬˮӡЧ��
			if SiteConfig("WatermarkOption")="Persits.Jpeg" then
				JpegPersits
			end if
		end if
		
		Jpeg.Save Server.MapPath(""&SaveFile&"")
		Set Jpeg = nothing
	end if
End Sub


'''''''''''''''''''''д��� Start'''''''''''''''''''
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

	Script=Script&"Bbsxp_InsertIntoEdit('"&SafeLongValue(Rs("UpFileID"))&"','"&SafeJsString(FileName)&"','"&SafeJsString(FileMIME)&"','"&SafeJsString(SaveFile)&"');"
	Rs.close
End Sub
'''''''''''''''''''''д���  End'''''''''''''''''''
%>