<%


class Csv_File

	Private filepath
	Private content
	Private lineArray
	Private separator
	
	Public Sub Class_Initialize()
		separator = vbCRLF
	End Sub
	
	Public Property Let CSVFile(f)
		filepath = f
		Call parse()
	End Property
	
	Public Property Get LineCount
		LineCount = UBound(lineArray)
	End Property
	
	Public Default Property Get Lines(i)
		t = lineArray(i)
		Set lin = new CSV_Line
		lin.lineString = t
		Set Lines = lin
	End Property
	
	Public Function parse()
		Set objFs = Server.CreateObject("Scripting.FileSystemObject")
		Set objTxt = objFs.OpenTextFile(filepath)
		content = objTxt.ReadAll
		lineArray = split(content,separator)
	End Function


end class


%>