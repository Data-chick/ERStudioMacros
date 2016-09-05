

Sub Main
	'initialize ER variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel


	EntCount = 0
	EntProbCount = 0


    Debug.Clear
	Debug.Print "ColumnNamesWithSpaces"

	'loop through each entity in the active model
	For Each MyEntity In MyModel.Entities
		For Each MyAttribute In MyEntity.Attributes


			If InStr(1,MyAttribute.ColumnName, " ") > 0 Then

				Debug.Print  MyEntity.TableName & " " & MyAttribute.ColumnName

			EntProbCount = EntProbCount +1
			End If



	EntCount = EntCount + 1
	Next MyAttribute

	Next MyEntity

Debug.Print EntProbCount & " Entities found out of " & EntCount & " Total Entities with problems"
	Wait (100)
End Sub
