<%
Public Const KB = 1024
Public Const MB = 1024

class Trophy_Logger

	Private MESSAGE_SEPARATOR
	Private TYPE_SEPARATOR
	Private name
	Private directory
	Private filename
	Private path
	Private maxfileSize
	Private fileStreamobject
	Private fileObject
	Private logFormat
	
	Public Property Let loggerName(n)
		name = n
	End Property
	
	Public Property Let logfilePath(p)
		directory = getFileDirectory(p)
		filename = getFileName(p)
		path = p
	End Property
	
	Public Property Let maxSize(m)
		if InStr(m,"K") > 0 then
			sizeInt = Mid(m,1,InStr(m,"K")-1)
			maxfileSize = KB*sizeInt
		elseif InStr(m,"M") > 0 then
			sizeInt = Mid(m,1,InStr(m,"K")-1)
			maxfileSize = MB*KB*sizeInt
		else
			maxfileSize = KB*m
		end if
	End Property
	
	Public Property Let format(f)
		logFormat = getDateString(Now,f)
		logFormat = replace(logFormat,"[S]",Session.SessionID)
	End Property
	
	Public Sub Class_Initialize()
		MESSAGE_SEPARATOR = vbCRLF
		TYPE_SEPARATOR = " / "
	End Sub
	
	Public Sub Class_Terminate()
		Set fileStreamObject = Nothing
		Set fileObject = Nothing
	End Sub
	
	Private Function prepare(m)
		prepare = replace(logFormat,"[M]",m)	
	End Function
	
	Private Function openFile()
		if Not objFs.FileExists(path) then
			Call objFs.CreateTextFile(path)
		end if
		Set fileStreamobject = objFs.OpenTextFile(path,8)
		Set fileObject = objFs.GetFile(path)
	End Function
	
	Private Function write(s)
		if Not IsObject(fileobject) then
			Call openFile()
		end if
		if fileobject.size > maxFileSize then
			Call backupFile()
		end if
		On Error Resume Next
		Call fileStreamobject.Write(s & MESSAGE_SEPARATOR)
		On Error Goto 0
	End Function
	
	Public Function closeFile()
		if isObject(fileStreamobject) then
			fileStreamobject.Close()
		end if
		Set fileObject = Nothing
	End Function
	
	Public Function append(str)
		Call write(prepare(str))
	End Function
	
	Public Function appendMessage(messageType,message)
		Call write(prepare(messageType & TYPE_SEPARATOR & message))
	End Function
	
	Private Function backupFile()
		Dim a
		a = true
		z = 1
		Do While a=true
			nfilename = directory & getFilePrefix(filename) & "_" & z & "." & getFileExtension(filename)
			if Not objFs.FileExists(nfilename) then
				Call objFs.CopyFile(path,nfilename)
				Call closeFile()
				Call objFs.DeleteFile(path)
				Call objFs.CreateTextFile(path)
				Call openFile()
				a = false
			end if
			z = z+1
		Loop
	End Function
	
end class


%>