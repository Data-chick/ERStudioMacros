'TITLE:  NAME FOREIGN CONSTRAINTS
'DESCRIPTION:  This macro will prompt the user with a dialog to
'specify the naming convention for Foreign Constraints.  It will
'then name  the constraints using the parent and child table names.
'It will also make sure the name is unique by adding an index to the
'end of the constraint name.
'
'*************************************************************
'*     This macro will not rename existing FK names!!!
'*************************************************************
'
'INPUT:
'	MAXIMUM CHARACTERS:  This is the max limit of the number of characters
'	For the entire constraint Name.  Not required.  Can be specified even if
' 	truncation option is None.
'	TRUNCATE NUMBER:  This limits the number of characters in both the parent and
'	child part of the relationship to the number specified.
'	PREFIX:  This is the prefix for the relationship name
'	SEPARATOR:  This goes between the parent and child part of the relatioship name.
'	SUFFIX:  This is the suffix for the relationship name.

' Updated by: Karen Lopez.  Fixed logic on appended numbers for duplicates, plus
' added separator before duplicate numbers.  Also added protection for existing names
'
'CREATE DATE:  11/19/2000
'LAST UPDATE:  12/15/2009



Sub Main
	Dim ParTable As String
	Dim chiTable As String
	Dim separator As String
	Dim prefix As String
	Dim sufix As String
	Dim mdl As Model
	Dim diag As Diagram
	Dim truncation As Integer
	Dim maxchs As Integer
	Dim relation As Relationship
	Dim RName As String
	Dim C As Integer
	Dim ind As Integer
	Dim suf As String
	Dim RelCnt As Integer
	Dim newind As Integer


	Debug.Clear

	' set er variables
	Set diag = DiagramManager.ActiveDiagram
	Set mdl = diag.ActiveModel


	Begin Dialog UserDialog 660,350,"TravelPort UP FK Constraint Naming Macro" ' %GRID:10,7,1,1
		Text 50,21,90,14,"Order:",.Text1
		CheckBox 150,203,90,14,"P&refix:",.chbxPrefix
		CheckBox 150,259,90,14,"&Suffix:",.chbxSuffix
		CheckBox 150,231,100,14,"S&eparator:",.chbxseparator
		Text 50,84,160,14,"Truncation:",.Text2
		Text 50,182,90,14,"Additions:",.Text3
		OptionGroup .Order
			OptionButton 270,35,160,14,"&Parent Then Child",.OptionButton1
			OptionButton 270,56,140,14,"&Child Then Parent",.OptionButton2
			OptionButton 130,35,110,14,"Parent Only",.OptionButton5
			OptionButton 130,56,90,14,"Child Only",.OptionButton6
		OptionGroup .Trunc
			OptionButton 80,140,90,14,"&None",.OptionButton3
			OptionButton 80,161,400,14,"&Use only this many characters of the Parent or Child name:",.OptionButton4
		TextBox 490,154,90,21,.numchars
		TextBox 310,196,90,21,.prefix
		TextBox 310,224,90,21,.separator
		TextBox 310,252,90,21,.suffix
		OKButton 330,301,120,28
		CancelButton 470,301,120,28
		Text 160,98,260,14,"Max characters for entire FK name:",.Text5
		TextBox 490,112,90,21,.MaxChars
		Text 160,119,310,14,"(Can be specified even when checking ""None"".)",.Text4
	End Dialog
	Dim dlg As UserDialog


	If Dialog(dlg) = -1 Then

		'declare an array to check for duplictate relations
		C = mdl.Relationships.Count
		ReDim Rships(1 To C) As String
		Dim rels As Variant
		rels = Rships
		ind = 1

		For Each Ent In mdl.Entities


		    Debug.Print "______________________________________________________________"
			Debug.Print " "
			Debug.Print ent.TableName & "
		    Debug.Print "______________________________________________________________"
			Debug.Print " "

		ind = 1



		For Each relation In ent.ChildRelationships



		If Len(relation.Name) = 0 Then

			ParTable = relation.ParentEntity.TableName
			chiTable = relation.ChildEntity.TableName

			verb = relation.VerbPhrase


			Debug.Print "     " & ParTable & "  " & verb & "  " & chiTable

			Debug.Print " "

			'Remove spaces from table names
			ParTable = Replace(ParTable," ","")
			chiTable = Replace(chiTable," ","")


			If dlg.trunc = 1 And dlg.numchars <> Empty Then	'truncate Left

				'truncate names of tables
				truncate = Abs(CInt(dlg.numchars))
				ParTable = Left(ParTable, truncate)
				chiTable = Left(chiTable, truncate)
			
			End If

			prefix = dlg.prefix
			suffix = dlg.suffix
			separator = dlg.separator

			'choose order, O = parent first, 1 = child first, 2 = parent only, 3 = child only
			If dlg.Order = 0 Then
				
				'Additions
				RName = ""
				'Add prefix if option is checked
				If dlg.chbxprefix = 1 Then
					RName = prefix
				End If

				RName = RName & ParTable

				'Add separator if option is checked
				If dlg.chbxseparator Then
					RName = RName & separator
				End If

				RName = RName & chiTable

				'Add suffix if option is checked
				If dlg.chbxsuffix Then
					RName = RName & suffix
				End If

			ElseIf dlg.order = 1 Then

				'Additions
				RName = ""

				'Add prefix if option is checked
				If dlg.chbxprefix = 1 Then
					RName = prefix
				End If

				RName = RName & chiTable

				'Add separator if option is checked
				If dlg.chbxseparator = 1 Then
					RName = RName & separator
				End If

				RName = RName & ParTable

				'Add suffix if option is checked
				If dlg.chbxsuffix = 1 Then
					RName = RName & suffix
				End If

			ElseIf dlg.order = 2 Then

				RName = ""

				'Add prefix if option is checked
				If dlg.chbxprefix = 1 Then
					RName = prefix
				End If

				RName = RName & ParTable

				'Add prefix if option is checked
				If dlg.chbxsuffix = 1 Then
					RName = RName & suffix
				End If

			Else 'dlg.order = 3

				RName = ""

				'add prefix if option is checked
				If dlg.chbxprefix = 1 Then
					RName = prefix
				End If

				RName = RName & chiTable

				'add suffix if option is checked
				If dlg.chbxsuffix = 1 Then
					RName = RName & suffix
				End If

			End If

			'truncate Name to max chars specified
			If dlg.maxchars <> Empty Then
				maxchs = Abs(CInt(dlg.maxchars))
				RName = Left(RName, maxchs)
			End If

			' Appends a number for each child relationship for this entity.
			' Debug.Print "Ind = " & ind

			If ind > 1 Then

				'add unique index to last of the relationship name
				suf = Trim(Str(ind))

				If Len(suf) = 1 Then
					suf = separator & "0" & suf
					'MsgBox (suf) For debugging only
				Else
					suf = separator & suf
				End If


			Else
				' First rel name, so append a 01 to end
				suf = separator & "01"
			End If

			'Check for max total length
			If dlg.maxchars <> Empty Then
					maxchs = Abs(CInt(dlg.maxchars))
					RName = Left(RName, maxchs)
					RName = Left(RName,(Len(RName) - Len(suf)))
			End If
			'	RName = Left(RName,(Len(RName) - Len(suf))) changed behaviour
			predupername = RName
			Debug.Print "        PreDupeName: " & Predupername
			RName = RName & suf
			Debug.Print "        Rname: " & RName
			newind = ind

			Debug.Print checkdups(rels, RName, ind)

			Do While checkdups(rels, RName, ind)
				newind = newind + 1
				suf = Trim(Str(newind))
				If Len(suf) = 1 Then
					suf = separator & "0" & suf

				Else
					suf = separator & suf
				End If


				RName = PreduperName & suf
				Debug.Print "****trying****  " & RName
			Loop






			RName = Replace(RName," ","")
			Debug.Print "         *******************************************************"
			Debug.Print "         *  New FK constraint:  " & RName
			Debug.Print "         *******************************************************"
			Debug.Print " "
			relation.Name = RName
			rels(ind) = RName
			ind = ind + 1

		Else
			'Debug.Print "     FK name already exists:  " & relation.Name
			'Debug.Print " "
			rels(ind) = relation.Name
			ind = ind + 1
		End If
		Next relation
	Next Ent

	End If ' dialog


End Sub


Function checkdups ( R As Variant, RelName As String, C As Integer ) As Boolean

	Dim flag As Boolean

	flag = False

	For i = 1 To C
		If R(i) = RelName Then
			flag = True
		End If
	Next i

	checkdups = flag
End Function
