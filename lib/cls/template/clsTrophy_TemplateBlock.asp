<%


class Trophy_TemplateBlock

	Private blockId
	Private blockType
	Private blockHeader
	Private blockBody
	Private blockFooter
	
	Public Property Get Id
		Id = blockId
	End Property
	
	Public Property Get bType
		bType = blockType
	End Property
	
	Public Property Get Header
		Header = blockHeader
	End Property
	
	Public Property Get Body
		Body = blockBody
	End Property
	
	Public Property Let Body(b)
		blockBody = b
	End Property
	
	Public Property Get Footer
		Footer = blockFooter
	End Property
	
	Public Default Property Get Content
		Content = blockHeader & blockBody & blockFooter
	End Property
	
	Public Sub Class_Initialize()
		'--
	End Sub
	
	Public Sub loadBlock(id,btype,header,body,footer)
		blockId = id
		blockType = btype
		blockHeader = header
		blockBody = body
		blockFooter = footer
	End Sub
	
end class
%>