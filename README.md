<div align="center">

## Select List Thingy


</div>

### Description

*UPDATED 10-23-00* Now with more options, including multiple select, size, and default selected option. Quickly and easily creates select list combo boxes from a database for your web forms. Uses .GetString rather than recordset looping.
 
### More Info
 
See Code Comments...

I place the code in an include file which I link to from the pages in which I am building select lists. I use this with SQL Server, not sure what results will be with other databases, but it should work okay.

Returns a string with all of the code for your select list. All you need to do to display it is a Response.Write() or <%=varname%>

Your Recordset can only contain two columns for this to work properly. Example SQL: "SELECT usr_id,usr_name FROM users". The first column will be used for the VALUE of the option, the second column will be the display value.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Will Knoll](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/will-knoll.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/will-knoll-select-list-thingy__4-6339/archive/master.zip)





### Source Code

```
<%
'***************************************
'* Function Name: buildCombo
'* Syntax : myVar = buildCombo
'* Parameters :
'*		rsRecordset - Name of an OPEN
'*					recordset with 2 fields
'*		strComboName - String with the name
'*					you want to appear in
'*					<select name='' id=''>
'*		strFirstOption - String with what
'*					the first option should
'*					be, ex:
'*			<option value="0">N/A</option>
'*		blnMultiple - Boolean for mult.
'*					option or not...should
'*					be FALSE unless you set
'*					strSize
'*		strSize - String for size (rows)
'*					to display
'*		strSelected - String of the
'*					selected value, if you
'*					want an option
'*					autoselected
'*
'* Description : Quickly creates a drop-
'*		down list without looping a
'*		recordset, and minimal
'*		concatination
'*
'* Example :
'*		myCombo = buildCombo(rsTmp,
'*		"usr_id","<option value='0'>N/A
'*		</option>",True,"5","0")
'*
'* Result :	Produces a 5 row multiple
'*		select list called usr_id with
'*		the first display as N/A and
'*		selected as default
'***************************************
function buildCombo(byref rsRecordset,strComboName,strFirstOption,blnMultiple,strSize,strSelected)
	Dim strCombo
	Dim strLstSize
	Dim i
	If Len(Trim(strSize)) > 0 Then
		strLstSize = " size=""" & strSize & """"
	Else
		strLstSize = ""
	End If
	If blnMultiple Then
		strCombo = "<select id=""" & strComboName & """ name=""" & strComboName & """ multiple" & strLstSize & ">"
	Else
		strCombo = "<select id=""" & strComboName & """ name=""" & strComboName & """" & strLstSize & ">"
	End If
	If strFirstOption <> "" Then
		strCombo = strCombo & strFirstOption & vbCrLf
	End If
	With rsRecordset
		If Not (.EOF And .BOF) Then
			If Not .BOF Then
				'Just making sure in case you were previously doing something else
				.MoveFirst
			End If
			strCombo = strCombo & "<option value='" & .GetString(,,"'>","</option>" & vbCrLf & vbTab & "<option value='","" )
		Else
			'No records returned
			buildCombo = "<select><option>No Records Returned</option></select>"
			exit function
		End If
	End With
	strCombo = strCombo & "</select>"
	i = InStrRev(strCombo,"<option value='</select>")
	If i > 0 Then
		'Remove the extra <option value=' from the end of the string
		'It will work fine with it in there, but we like to be tidy don't we?
		strCombo = Replace(strCombo,"<option value='</select>","</select>")
	End If
	If Len(Trim(strSelected)) > 0 Then
		strCombo = Replace(strCombo,"<option value='" & strSelected & "'>","<option value='" & strSelected & "' selected>",1,1)
	End If
	buildCombo = strCombo
end function
%>
```

