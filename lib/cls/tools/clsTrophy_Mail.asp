<%

class Trophy_Mail

    Private mailerType
	Private mailerObjectNames
	Private attachments
    Private mailFormat
    Private mailBody
    Private mailSubject
    Private mailFrom
    Private reciepents
    Private mailServer
	Private mailResult
    Public mailerObject
    
	Public Property Get mType
		mType = mailerType
	End Property
	
	Public Property Let mType(t)
		mailerType = t
	End Property
	
	Public Property Get mServer
		mServer = mailServer
	End Property
	
	Public Property Let mServer(s)
		mailServer = s
	End Property
	
	Public Property Get format
		format = mailFormat
	End Property
	
	Public Property Let format(f)
		if inArray(array("html","plain"),f) then
			mailFormat = f
		else
			Call objError.trigger("trophy_email","wrong email format. only html and plain are supported")
		end if
	End Property
	
	Public Property Get from
		from = mailFrom
	End Property
	
	Public Property Let from(f)
		mailFrom = f
	End Property
	
	Public Property Let subject(s)
		mailSubject = s
	End Property
	
	Public Property Get subject
		subject = mailSubject
	End Property
	
	Public Property Get body
		body = mailBody
	End Property
	
	Public Property Let body(b)
		mailBody = b
	End Property
	
	Public Property Get result
		result = mailResult
	End Property
	
	Private Function loadDefaults()
		if isObject(objTrophy) then
			if objTrophy.config.parameters.exists("default-mail-component") then
				mailerType = objTrophy.config.parameters.Item("default-mail-component")
			end if
			if objTrophy.config.parameters.exists("default-mail-server") then
				mailServer = objTrophy.config.parameters.Item("default-mail-server")
			end if
		end if
	End Function
	
	Private Function addReciepentMail(m)
		if Not reciepents.Exists(m) then
			Call reciepents.Add(m,reciepents.Count)
		end if
	End Function
	
	Private Function send_JMail() 
		mailerObject.ServerAddress = mailServer
		mailerObject.Sender = From
		For Each email in reciepents
			mailerObject.AddRecipient email
		Next
		mailerObject.subject = subject
		if format = "html" then
			mailerObject.HTMLBody = Body
		else
			mailerObject.Body = Body 
		end if
		For Each file in attachments
			Call mailerObject.addAttachment(file)
		Next
		mailerObject.Execute
    End Function
	
	Private Function send_CDonts()
		For Each email in reciepents
			mailerObject.To email
		Next
		mailerObject.From = From
		mailerObject.Body = Body
		mailerObject.Subject = Subject
        if Format = "html" then
            mailerObject.BodyFormat = 0
            mailerObject.MailFormat = 0
        end if
		For Each file in attachments
			Call mailerObject.attachFile(file)
		Next
        mailerObject.Send
	End Function
	
	Private Function send_ASPEmail()
		mailerObject.Host = MailServer	
		mailerObject.From = From
		For Each email in reciepents
			mailerObject.AddAddress email
		Next
		mailerObject.AddAddress Email
		mailerObject.Subject = Subject
		mailerObject.Body = Subject
		 if Format = "html" then
            mailerObject.IsHTML = True
        else
		    mailerObject.IsHTML = False
       	end if
		For Each file in attachments
			Call mailerObject.addAttachment(file)
		Next
		result = mailer.Send()
	End Function
	
	Private Function loadInstalledComponent()
		Err.Clear
		On Error Resume Next
		For Each objectName in mailerObjectNames
			Set controlObject = Server.CreateObject(mailerObjectNames.Item(objectName))
			if Err.Number = 0 then
				mailerType = objectName
				Exit For
			end if
		Next 
		On Error GoTo 0
	End Function
	
	Private Function setDefaultObjectNames()
		Call mailerObjectNames.Add("JMail","JMail.SMTPMail")
		Call mailerObjectNames.Add("CDonts","CDONTS.NewMail")
		Call mailerObjectNames.Add("ASPEmail","Persits.MailSender")
	End Function
	
    Public Sub Class_Initialize()
        format = "html" 
		Set attachments = Server.CreateObject("Scripting.Dictionary")
		Set reciepents = Server.CreateObject("Scripting.Dictionary")
		Set mailerObjectNames = Server.CreateObject("Scripting.Dictionary")
		Call loadDefaults()
		Call setDefaultObjectNames()
    End Sub
    
	Public Function addReciepent(email)
		Call addReciepentMail(email)
	End Function
	
	Public Function addAttachment(f)
		Call attachments.Add(f,attachments.Count)
	End Function
	
    Public Function run()
		Dim mailCmd
		if mailerType = "" then
			Call loadInstalledComponent()
		end if
		Set mailerObject = Server.CreateObject(mailerObjectNames.Item(mailerType))
		mailCmd = "Call send_" & mailerType & "()"
		Call Execute(mailCmd)
    End Function
	
	Public Function send(f,s,r,b)
		from = f
		subject = s
		if isArray(r) then
			reciepents = r
		else
			Call addReciepent(r)
		end if
		body = b
		Call run()
	End Function
	
	

end Class


%>
