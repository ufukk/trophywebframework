
[!container type="form" id="frm1" name="form1" action="controller.asp?action=test5" method="POST" validationfile="form1-validator.xml"]
	[!element type="label" for="field_3" value="Field3"]
	<br>
	[!element type="field" ftype="textfield" name="field_3" value="dasddaa"]
	<br>
	Field4:
	<br>
	[!element type="field" ftype="textfield" name="field_4" value="4"]
	<br>
	Field5:
	<br>
	[!element type="field" ftype="textfield" name="field_5" value="val" maxlength="10"]
	<br>
	Field6:
	<br>
	[!element type="field" ftype="textfield" name="field_6" value="val" maxlength="10"]
	<br>
	<br><br>
	<input type="submit" value=" GÖNDER ">
[!/container]
[!include file="f:trophy-form-validation.tpl"]
