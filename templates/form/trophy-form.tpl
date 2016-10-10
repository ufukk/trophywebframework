
[!if condition="{form .validate}"]
	<script language="javascript">
		function minimumLength(val,minimum) {
			if(val.value.length < minimum) {
				return false;
			} else {
				return true;
			}
		}
		
		function lengthBetween(val,minimum,maximum) {
			if(val.value.length > minimum && val.value.length < maximum) {
				return true;
			} else {
				return false;
			}
		}
		
		function regex(val,pattern) {
			eval("var re = /"+pattern+"/g;")
			return re.test(val);
		}
		
		function isValidEmail(val) {
			return regex(val,"[a-zA-Z0-9_.-]{1,255}@[a-zA-Z0-9_.-]{1,255}.[a-zA-Z0-9_.-]{1,10}");
		}
		
		function isNumeric(val) {
			return (parseInt(val));
		} 
		
		function isFloat(val) {
			return (parseFloat(val));
		} 
	
		function cF_[!form .name](frm) {
			[!loop name="VALIDATE" var="fields"] [!if condition=`{$VALIDATE .validate}`]
				if (![!$VALIDATE .validation]) {
					alert('[!$VALIDATE .title] boþluðunu doldurmalýsýnýz');
					frm.[!$VALIDATE .name].focus();
					return false;
				}
				[!/if] [!/loop]
			return true;
		}
	</script>
[!/if]

<table border="0" cellpadding="0" cellspacing="0">
<form action="[!form .action]" name="[!form .name]" method="[!form .method]" onSubmit="[!form .onSubmit]"[!ENCTYPE_STR]>	<tr>
		<td>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td class="text1" align="left">
			[!FIELDS]
		</td>
	</tr>
	<tr>
		<td>
			<input type="submit" value="[!form .submitLabel]">
		</td>
	</tr>
	</form>
</table>
