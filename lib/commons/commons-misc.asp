<%
Public Function echo(str)
	Response.Write str 
End Function

Public Function debugPrint(s)
	response.write "<PRE>" & s & "</PRE>"
End Function

Public Function printLine(str)
	echo str & "<br>"
End Function

Public Function debugMessage(ByRef obj,message)
	echo "DEBUG MESSAGE: " & TypeName(obj) & ": " & message & "<br>"
End Function

Public Function inArray(ByRef arr,ByVal key)
	Dim r
	r = false
	For a=0 to UBound(arr) 
		if arr(a) = key then
			r = true
		end if
	Next
	inArray = r
End Function

Public Function reverseArray(arr)
	Dim a,r()
	a = false
	For x = UBound(arr) to 0 Step -1
		if a = false then
			Redim Preserve r(1)
			r(0) = arr(x)
			a = True
		else
			Redim Preserve r(UBound(r))
			r(UBound(r)) = arr(x)
		end if
	Next
	reverseArray = r
End Function

Public Function pushArray(arr,element)
	if Not UBound(arr) then
		Redim Preserve arr(UBound(arr)+1)
		arr(UBound(arr)) = element
	else
		arr(0) = element
	end if
	pushArray = arr
End Function

 Public Function deleteArray(arr,pos)
     Dim index
     For index = pos to UBound(arr) - 1
       arr(index) = arr(index + 1)
     Next
	Redim Preserve arr(UBound(arr) - 1)
	deleteArray = arr
End Function

Public Function readFile(path) 
	Set objFs = Server.CreateObject("Scripting.FileSystemObject")
	Set objF = objFs.OpenTextFile(path)
	output = objF.ReadAll
	Set objFs = Nothing
	Set objF = Nothing
	readFile = output
end function

Public Function getRandom(max)
	Randomize
	rC = Int(max*Rnd)
	getRandom = rC
End Function

Public Function newDictionary()
	Set newDictionary = Server.CreateObject("Scripting.Dictionary")
End Function

Public Function getDictionary(keys,values)
	Set nw = newDictionary()
	if isArray(keys) then
		For i=0 to UBound(keys)
			nw.Item(keys(i)) = values(i)
		Next
	end if
	Set getDictionary = nw
End Function

Public Function newFileSystem()
	Set newFileSystem = Server.CreateObject("Scripting.FileSystemObject")
End Function

Public Function newXMLParser()
	Set newXMLParser = Server.CreateObject("Microsoft.XMLDOM")
End Function

Public Function newRegexp()
	Dim regexpObj
	Set regexpObj = new Regexp
	regexpObj.Global = True
	regexpObj.Multiline = True
	regexpObj.IgnoreCase = True
	Set newRegexp = regexpObj
End Function

Public Function generateID(slen,useinteger,uselowercase,useuppercase) 
	chars = ""
	if useinteger then
		chars = chars & "0123456789"
	end if
	if uselowercase then
		chars = chars & "abcdefghijklmnoprqstvywxz"
	end if
	if useuppercase then
		chars = chars & "ABCDEFGHIJKLMNOPQRSTVWXYZ"
	end if
	lN = Len(chars)-1
	output = ""
	Randomize
	For x = 0 to slen-1
		rC = Int(lN*Rnd)
		if rC = 0 then
			rC = 1
		end if
		output = output & Mid(chars,rC,1)	
	Next
	generateID = output
End Function


Public Function isInteger(s)
	Set r = newRegexp
	r.Pattern = "([^0-9])"
	Set results = r.Execute(s)
	if results.Count = 0 then
		isInteger = True
		Exit Function
	else
		isInteger = False
		Exit Function
	end if
End Function

Public Function redirect(url)
	Response.Redirect url
End Function



%>