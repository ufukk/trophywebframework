<%


class Trophy_UserObject

	Private objectName
	Private objectClassName
	Private objectDataSource
	Private objectDBObjectName
	
	Public Property Get name
		name = objectName
	End Property
	
	Public Property Let name(n)
		objectName = n
	End Property
	
	Public Property Get className
		className = objectClassName
	End Property
	
	Public Property Let className(n)
		objectClassName = n
	End Property
	
	Public Property Get dataSource
		dataSource = objectDataSource
	End Property
	
	Public Property Let dataSource(n)
		objectDataSource = n
	End Property
	
	Public Property Get dbObjectName
		dbObjectName = objectDBObjectName
	End Property
	
	Public Property Let dbObjectName(n)
		objectDBObjectName = n
	End Property
	
end class

%>