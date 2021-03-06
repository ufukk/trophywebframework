<%
' ------------------------------------------------------------------------------
'	Author:		Lewis Moten
'	Email:		Lewis@Moten.com
'	URL:		http://www.lewismoten.com
'	Date:		March 19, 2002
' ------------------------------------------------------------------------------

' Field class represents interface to data passed within one field
'
' Two available methods of getting a field:
'	Set objField = objUpload.Fields("File1")
'	Set objField = objUpload("File1")
'
'
'	objField.Name
'		Name of the field as defined on the form
'
'	objFiled.Filepath
'		Path that file was sent from
'
'		ie: C:\Documents and Settings\lmoten\Desktop\Photo.gif
'
'	objField.FileDir
'		Directory that file was sent from
'
'		ie: C:\Documents and Settings\lmoten\Desktop
'
'	objField.FileExt
'		Uppercase Extension of the file
'
'		ie: GIF
'
'	objField.FileName
'		Name of the file
'
'		use: Response.AddHeader "Content-Disposition", "filename=""" & objField.FileName & """"
'
'		ie: Photo.gif
'
'	objField.ContentType
'		Type of binary data
'
'		use: Response.ContentType = objField.ContentType
'
'		ie: image/gif
'
'	objField.Value
'		Unicode value passed from form.  This value is empty if the field is binary data.
'
'		use: Response.Write "The value of this field is: " & objField.Value
'
'	objField.BinaryData
'		Contents of files binary data. (Integer SubType Array)
'
'		use: Response.BinaryWrite objField.BinaryData
'
'	objField.BLOB
'		Same thing as BinaryData but with a shorter name.  Added to help prevent
'		confusion with database access.
'
'		use: Call lobjRs.Fields("Image").AppendChunk(objField.BLOB)
'
'	objField.Length
'		byte size of Value or BinaryData - depending on type of field
'
'		use: Response.Write "The size of this file is: " & objField.Length
'
'	objField.BinaryAsText()
'		Converts binary data into unicode format.  Useful when you expect the user
'		to upload a text file and you have the need to interact with it.
'
'		use: Response.Write objField.BinaryAsText()
'
'	objField.SaveAs()
'		Saves binary data to a specified path.  This will overwrite any existing files.
'
'		use: objField.SaveAs(Server.MapPath("/Uploads/") & "\" & objField.FileName)
'
' ------------------------------------------------------------------------------
Class Trophy_RequestBinaryField
	
	Public Name				' Name of the field defined in form

	Private mstrPath		' Full path to file on visitors computer
							' C:\Documents and Settings\lmoten\Desktop\Photo.gif
	
	Public FileDir			' Directory that file existed in on visitors computer
							' C:\Documents and Settings\lmoten\Desktop

	Public FileExt			' Extension of the file
							' GIF

	Public FileName			' Name of the file
							' Photo.gif
	
	Public ContentType		' Content / Mime type of file
							' image/gif
							
	Public Value			' Unicode value of field (used for normail form fields - not files)
	
	Public BinaryData		' Binary data passed with field (for files)

	Public Length			' byte size of value or binary data

	Private mstrText		' Text buffer 
								' If text format of binary data is requested more then
								' once, this value will be read to prevent extra processing
	
' ------------------------------------------------------------------------------
	Public Property Get BLOB()
		BLOB = BinaryData
	End Property
' ------------------------------------------------------------------------------
	Public Function BinaryAsText()
		
		' Binary As Text returns the unicode equivilant of the binary data.
		' this is useful if you expect a visitor to upload a text file that
		' you will need to work with.
		
		' NOTICE:
		' NULL values will prematurely terminate your Unicode string.
		' NULLs are usually found within binary files more often then plain-text files.
		' a simple way around this may consist of replacing null values with another character
		' such as a space " "

		Dim lbinBytes
		Dim lobjRs
		
		' Don't convert binary data that does not exist
		If Length = 0 Then Exit Function
		If LenB(BinaryData) = 0 Then Exit Function
		
		' If we previously converted binary to text, return the buffered content
		If Not Len(mstrText) = 0 Then
			BinaryAsText = mstrText
			Exit Function
		End If
		
		' Convert Integer Subtype Array to Byte Subtype Array
		lbinBytes = ASCII2Bytes(BinaryData)
   		
   		' Convert Byte Subtype Array to Unicode String
   		mstrText = Bytes2Unicode(lbinBytes)
   		
   		' Return Unicode Text
    	BinaryAsText = mstrText

	End Function
' ------------------------------------------------------------------------------
	Public Sub SaveAs(ByRef pstrFileName)

		Dim lobjStream
		Dim lobjRs
		Dim lbinBytes
		
		' Don't save files that do not posess binary data
		If Length = 0 Then Exit Sub
		If LenB(BinaryData) = 0 Then Exit Sub
		
		' Create magical objects from never never land
		Set lobjStream = Server.CreateObject("ADODB.Stream")
		
		' Let stream know we are working with binary data
		lobjStream.Type = adTypeBinary
		
		' Open stream
		Call lobjStream.Open()

		' Convert Integer Subtype Array to Byte Subtype Array
		lbinBytes = ASCII2Bytes(BinaryData)
		
		' Write binary data to stream
		Call lobjStream.Write(lbinBytes)
		
		' Save the binary data to file system
		'	Overwrites file if previously exists!
		Call lobjStream.SaveToFile(pstrFileName, adSaveCreateOverWrite)
		
		' Close the stream object
		Call lobjStream.Close()
		
		' Release objects
		Set lobjStream = Nothing
	
	End Sub
' ------------------------------------------------------------------------------
	
	Public Default Property Get Value_()
		Value_ = Value
	End Property
	
	Public Property Let FilePath(ByRef pstrPath)
		
		mstrPath = pstrPath
		
		' Parse File Ext
		If Not InStrRev(pstrPath, ".") = 0 Then
			FileExt = Mid(pstrPath, InStrRev(pstrPath, ".") + 1)
			FileExt = UCase(FileExt)
		End If
		
		' Parse File Name
		If Not InStrRev(pstrPath, "\") = 0 Then
			FileName = Mid(pstrPath, InStrRev(pstrPath, "\") + 1)
		else
			'ufux: mozilla compability
			FileName = pstrPath
		End If
		
		' Parse File Dir
		If Not InStrRev(pstrPath, "\") = 0 Then
			FileDir = Mid(pstrPath, 1, InStrRev(pstrPath, "\") - 1)
		End If
		
	End Property
' ------------------------------------------------------------------------------
	Public Property Get FilePath()
		FilePath = mstrPath
	End Property
' ------------------------------------------------------------------------------
	Private Function ASCII2Bytes(ByRef pbinBinaryData)
	
		Dim lobjRs
		Dim llngLength
		Dim lbinBuffer
		
		' get number of bytes
		llngLength = LenB(pbinBinaryData)
		
		Set lobjRs = Server.CreateObject("ADODB.Recordset")
		
		' create field in an empty recordset to hold binary data
		Call lobjRs.Fields.Append("BinaryData", adLongVarBinary, llngLength)
		
		' Open recordset
		Call lobjRs.Open()
		
		' Add a new record to recordset
		Call lobjRs.AddNew()
		
		' Populate field with binary data
		Call lobjRs.Fields("BinaryData").AppendChunk(pbinBinaryData & ChrB(0))
		
		' Update / Convert Binary Data
			' Although the data we have is binary - it has still been
			' formatted as 4 bytes to represent each byte.  When we
			' update the recordset, the Integer Subtype Array that we
			' passed into the Recordset will be converted into a
			' Byte Subtype Array
		Call lobjRs.Update()
		
		' Request binary data and save to stream
		lbinBuffer = lobjRs.Fields("BinaryData").GetChunk(llngLength)
		
		' Close recordset
		Call lobjRs.Close()
		
		' Release recordset from memory
		Set lobjRs = Nothing
		
		' Return Bytes
		ASCII2Bytes = lbinBuffer
		
	End Function
' ------------------------------------------------------------------------------
	Private Function Bytes2Unicode(ByRef pbinBytes)

		Dim lobjRs
		Dim llngLength
		Dim lstrBuffer
		
		llngLength = LenB(pbinBytes)
				
		Set lobjRs = Server.CreateObject("ADODB.Recordset")

		' Create field in an empty recordset to hold binary data
    	Call lobjRs.Fields.Append("BinaryData", adLongVarChar, llngLength)
    	
    	' Open Recordset
    	Call lobjRs.Open()
    	
    	' Add a new record to recordset
    	Call lobjRs.AddNew()
    	
    	' Populate field with binary data
    	Call lobjRs.Fields("BinaryData").AppendChunk(pbinBytes)
    	
    	' Update / Convert.
    		' Ensure bytes are proper subtype
    	Call lobjRs.Update()
    	
    	' Request unicode value of binary data
    	lstrBuffer = lobjRs.Fields("BinaryData").Value
    	
    	' Close recordset
    	Call lobjRs.Close()

    	' Release recordset from memory
    	Set lobjRs = Nothing
	
		' Return Unicode
		Bytes2Unicode = lstrBuffer
		
	End Function

' ------------------------------------------------------------------------------
End Class
' ------------------------------------------------------------------------------
%>