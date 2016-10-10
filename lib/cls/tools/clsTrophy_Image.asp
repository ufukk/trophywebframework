<%
'An object oriented approach to Mike Shaffer's image library. 

class Trophy_Image

	Private imagePath
	Private originalImagePath
	Private imageWidth
	Private imageHeight
	Private imageType
	Private imageFileSize
	Private imageDepth
	Private loadError
	Private errorNumber
	
	Public Property Get width
		width = imageWidth 
	End Property
	
	Private Property Let width(w)
		imageWidth = w
	End Property
	
	Public Property Get height
		height = imageHeight
	End Property
	
	Private Property Let height(h)
		imageHeight = h
	End Property
	
	Public Property Get iType
		iType = imageType
	End Property
	
	Public Property Get fileSize
		size = imageFileSize
	End Property
	
	Public Property Get depth
		depth = imageDepth
	End Property
	
	Private Property Let depth(d)
		imageDepth = d
	End Property
	
	Public Property Get path
		path = imagePath
	End Property
	
	Public Property Get originalPath
		originalPath = originalImagePath
	End Property
	
	Public Property Let path(p)
		originalImagePath = p
		imagePath = Server.MapPath(p)
		Call loadImage()
	End Property
	
	Public Property Let absolutePath(p)
		originalImagePath = p
		imagePath = p
		Call loadImage()
	End Property
	
	Public Property Get error
		error = loadError
	End Property
	
	Public Property Get html(keys,values)
		html = "<img src=""" & originalPath & """ width=""" & width & """ height=""" & height & """"
		if isArray(keys) then
			For i=0 to UBound(keys)
				if i <> UBound(keys)-1 then
					html = html & " "
				end if
				html = html & keys(i) & "=" & """" & values(i) & """"
			Next
		end if
		html = html & ">"
	End Property
	
	Private Function setError()
		 loadError = true
	End Function
	
	Private Function clearError()
		loadError = false
	End Function
	
	Private Function GetBytes(offset, bytes)
		Dim objFSO
     	Dim objFTemp
     	Dim objTextStream
     	Dim lngSize
		'on error resume next
		if Not isObject(objFs) then
			Set objFs = CreateObject("Scripting.FileSystemObject")
     	end if
		' First, we get the filesize
     	Set objFTemp = objFs.GetFile(path)
     	lngSize = objFTemp.Size
     	set objFTemp = nothing
	 	fsoForReading = 1
     	Set objTextStream = objFs.OpenTextFile(path,fsoForReading)
	 	if offset > 0 then
        	strBuff = objTextStream.Read(offset - 1)
     	end if
	 	if bytes = -1 then		' Get All!
        	GetBytes = objTextStream.Read(lngSize)  'ReadAll
     	else
	        GetBytes = objTextStream.Read(bytes)
     	end if
     	objTextStream.Close
    	set objTextStream = nothing
  	End Function
	
	Private Function lngConvert(strTemp)
     lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
  	End Function

  	Private Function lngConvert2(strTemp)
    	 lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
  	End Function
	
	Private Function loadImage()
	 dim strPNG 
     dim strGIF
     dim strBMP
     dim strType
	 Call clearError()
     strType = ""
     strImageType = "(unknown)"
     gfxSpex = False
     strPNG = chr(137) & chr(80) & chr(78)
     strGIF = "GIF"
     strBMP = chr(66) & chr(77)
     strType = GetBytes(0, 3)
     if strType = strGIF then				' is GIF
        strImageType = "GIF"
        Width = lngConvert(GetBytes(flnm, 7, 2))
        Height = lngConvert(GetBytes(flnm, 9, 2))
        Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
        gfxSpex = True
     elseif left(strType, 2) = strBMP then		' is BMP
        strImageType = "BMP"
        Width = lngConvert(GetBytes(flnm, 19, 2))
        Height = lngConvert(GetBytes(flnm, 23, 2))
        Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
        gfxSpex = True
     elseif strType = strPNG then			' Is PNG
        strImageType = "PNG"
        Width = lngConvert2(GetBytes(flnm, 19, 2))
        Height = lngConvert2(GetBytes(flnm, 23, 2))
        Depth = getBytes(flnm, 25, 2)
        select case asc(right(Depth,1))
           case 0
              Depth = 2 ^ (asc(left(Depth, 1)))
              gfxSpex = True
           case 2
              Depth = 2 ^ (asc(left(Depth, 1)) * 3)
              gfxSpex = True
           case 3
              Depth = 2 ^ (asc(left(Depth, 1)))  '8
              gfxSpex = True
           case 4
              Depth = 2 ^ (asc(left(Depth, 1)) * 2)
              gfxSpex = True
           case 6
              Depth = 2 ^ (asc(left(Depth, 1)) * 4)
              gfxSpex = True
           case else
              Depth = -1
        end select
     else
        strBuff = GetBytes(0, -1)		' Get all bytes from file
        lngSize = len(strBuff)
        flgFound = 0
        strTarget = chr(255) & chr(216) & chr(255)
        flgFound = instr(strBuff, strTarget)
        if flgFound = 0 then
           exit function
        end if
        strImageType = "JPG"
        lngPos = flgFound + 2
        ExitLoop = false
        do while ExitLoop = False and lngPos < lngSize
           do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
              lngPos = lngPos + 1
           loop
           if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
              lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
              lngPos = lngPos + lngMarkerSize  + 1
           else
              ExitLoop = True
           end if
       loop
       if ExitLoop = False then
          Width = -1
          Height = -1
          Depth = -1
		  Call setError()
       else
          Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
          Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
          Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
          Call clearError()
       end if
     end if
  End Function
	
end class

%>