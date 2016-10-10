<!--#include file="clsTrophy_RequestBinaryHandler.asp"-->
<%
Public Const T_ASCII = 0
Public Const T_BINARY = 1
Public Const M_GET = -1
Public Const M_POST = -2
Public Const VALIDATOR_FIELD_PREFIX = "validator_"

class Trophy_RequestHandler
	
	Public RequestType
	Public RequestMethod
	Private filterLength
	Private filters
	Private userVariables
	Public binaryObject
	Private validatorMessages
	Private isPostback
	
	Public Sub Class_Initialize()
		if REQ_ASCII_MODE <> true then
			Call loadRequestData()
		else
			RequestType = T_ASCII
		end if
		if Request.ServerVariables("REQUEST_METHOD") = "POST" then
			RequestMethod = M_POST
		elseif Request.ServerVariables("REQUEST_METHOD") = "GET" then
			RequestMethod = M_GET
		end if
		Set userVariables = newDictionary()
		Set validatorMessages = newDictionary()
	End Sub	
	
	Public Property Let requestFilters(f)
		filters = f
	End Property
	
	Public Property Get IsGET
		Dim r
		if RequestMethod=M_GET then
			r = True 
		else
			r = False
		end if
		IsGet = r
	End Property
	
	Public Property Get IsPOST
		Dim r
		if RequestMethod=M_POST then
			r = True 
		else
			r = False
		end if
		IsPOST = r
	End Property
	
	Public Property Get URL
		Dim u,prt
		prt = LCase(Mid(Request.ServerVariables("SERVER_PROTOCOL"),1,InStr(Request.ServerVariables("SERVER_PROTOCOL"),"/")-1)) & "://"
		u = prt &  Request.ServerVariables("SERVER_NAME") & "" & Request.ServerVariables("SCRIPT_NAME")
		if Not IsNull(Request.ServerVariables("QUERY_STRING")) then
			u = u & "?" & Request.ServerVariables("QUERY_STRING")
		end if
		URL = u
	End Property
	
	Public Default Property Get field(f)
		Dim v
		if userVariables.Exists(f) then
			v = userVariables.Item(f)
		elseif Request.ServerVariables(CStr(f)) = "" then
			Select Case RequestType
				Case T_ASCII
					v = getAsciiValue(f) 
				Case T_BINARY
					v = getBinaryValue(f)
			End Select
			v = getFieldValue(f)
		else
			v = Request.ServerVariables(f)
		end if
		field = v
	End Property
	
	Public Property Let field(f,v)
		if userVariables.Exists(f) then
			userVariables.item(f) = v
		else
			Call userVariables.Add(f,v)
		end if
	End Property
	
	Public Property Get isNonValidated
		isNonValidated = isPostback
	End Property
	
	Public Function getValidatorMessages()
		Set getValidatorMessages = validatorMessages
	End Function
	
	Public Function getNonValidatedValue(fieldName)
		getNonValidatedValue = Me.Field(VALIDATOR_FIELD_PREFIX & fieldName)
	End Function
	
	Public Function getNonValidatedMessage(fieldName)
		getNonValidatedMessage = validatorMessages.Item(fieldName)
	End Function
	
	Public Function getFieldValue(key)
		Select Case RequestType
			Case T_ASCII
				v = "getAsciiValue(""" & key & """)" 
			Case T_BINARY
				v = "getBinaryValue(""" & key & """)"
		End Select
		getFieldValue = applyFilters(v)
	End Function
	
	Public Property Get valueMap(fieldName)
		valueMap = split(field(fieldName),",") 
	End Property
	
	Public Property Get Count
		Dim r
		Select Case RequestType
			Case T_ASCII
				Select Case RequestMethod
					Case M_GET
						r = Request.QueryString.Count
					Case M_POST
						r = Request.Form.Count
				End Select
			Case T_BINARY
				v = binaryObject.FieldCount
		End Select
		Count = r
	End Property
	
	Public Sub Class_Terminate()
	End Sub
	
	Public Function addFilter(funcname)
		Redim Preserve addFilter(filterLength)
		filters(filterLength) = funcname
		filterLength = filterLength+1
	End Function
	
	Private Sub loadRequestData()
		Dim total,byteSize,byteData
		filterLength = 1
		contentType = Request.ServerVariables("HTTP_CONTENT_TYPE")
		if InStr(contentType,"multipart") = 1 then
			isMultipart = true
			byteSize = Request.TotalBytes
			byteData = Request.BinaryRead(byteSize)
			total = CLng(Len(byteData))
			dataArray = getFromSession()
		else
			isMultipart = false
		end if
		if isNull(dataArray) And total > 0 then
			Call savetoSession(byteData,byteSize)
		end if
		if isMultipart then
			RequestType = T_BINARY
			Set binaryObject = new Trophy_RequestBinaryHandler
			Call binaryObject.parseRequestData(byteData,byteSize)
			if Not IsNull(dataArray) then
				Call binaryObject.parseRequestData(dataArray(0),dataArray(1))
			else
				Call binaryObject.parseRequestData(byteData,byteSize)
			end if
		else
			RequestType = T_ASCII
		end if
	End Sub

	Private Function getAsciiValue(k)
		getAsciiValue = Request(k)
	End Function
	
	Private Function getBinaryValue(k) 
		getBinaryValue = binaryObject.Fields(k)
	End Function
	
	Public Function isFile(n)
		if getFile(n).FileName <> "" then
			isFile = true
		else
			isFile = false
		end if
	End Function
	
	Public Function getFile(k)
		Set getFile = binaryObject.Fields(k)
	End Function
	
	Private Function applyFilters(cmd1)
		Dim c,filterCmd,result
		c = 0
		filterCmd = cmd1
		For c=0 to UBound(filters)
			filterCmd = filters(c) & "(" & filterCmd & ")"
		Next
		filterCmd = "result = " & filterCmd
		Call Execute(filterCmd)
		applyFilters = result
	End Function
	
	Private Function setAsciiValue(k,v)
		Request(k) = v
	End Function
	
	Private Function setBinaryValue(k,v) 
		binaryObject.Fields(k).Value = v
	End Function
	
	Public Function SaveFile(fieldName,path)
		if RequestType=T_BINARY then
			Call getFile(fieldName).SaveAs(path)
		end if
	End Function
	
	Public Property Get Iterator
		Dim o
		if RequestType=T_ASCII then
			if RequestMethod=M_POST then
				Set o = Request.Form
			else
				Set o = Request.QueryString
			end if
		elseif RequestType=T_BINARY then
			 Call binaryObject.loadIterator()
			 Set o = binaryObject.fieldDict
		end if
		Set Iterator = o
	End Property
	
	Public Function loadValidatorFields()
		if isObject(Session(CONFIG_SESSION_REQUEST_KEY)) then
			isPostback = true
			For Each item in Session(CONFIG_SESSION_REQUEST_KEY)
				itemName = VALIDATOR_FIELD_PREFIX & item
				Me.Field(itemName) = Session(CONFIG_SESSION_REQUEST_KEY).Item(item)
			Next
			For Each item in Session(CONFIG_SESSION_VALIDATOR_KEY)
				Call validatorMessages.Add(item,Session(CONFIG_SESSION_VALIDATOR_KEY).Item(item))
			Next
		end if
		Session.Contents.Remove(CONFIG_SESSION_VALIDATOR_KEY)
		Session.Contents.Remove(CONFIG_SESSION_REQUEST_KEY)
	End Function
	
	Private Function getFromSession()
		if LenB(Session("_tmp_byte_data")) > 0 then
			Dim r(2)
			r(0) = Session("_tmp_byte_data")
			r(1) = Session("_tmp_byte_size")
		else
			r = Null
		end if
		'Session.Contents.Remove("_tmp_byte_data")
		'Session.Contents.Remove("_tmp_byte_size")
		getFromSession = r
	End Function
	
	Private Function savetoSession(data,size)
		Session("_tmp_byte_data") = data
		Session("_tmp_byte_size") = size
	End Function
	
end class


%>