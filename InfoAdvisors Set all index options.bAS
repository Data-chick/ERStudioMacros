'Author: Karen Lopez
'www.infoadvisors.com
'Macro Purpose: Set *all* Index options to the DB2 default


Sub Main
	Dim theDiagram As Diagram
	Dim theModel As Model
	Dim theSubModel As SubModel
	Dim ent As Entity
	Dim indx As Index
	Dim indxname As String

	Debug.Clear

	Set theDiagram = DiagramManager.ActiveDiagram
	Set theModel = theDiagram.ActiveModel
	Debug.Print "Setting Index options"
	Debug.Print " "

	For Each ent In theModel.Entities
		Debug.Print ent.TableName

		For Each indx In ent.Indexes
			indxname = indx.Name
			Debug.Print "     " + indx.Name
			indx.MinPercentUsed = 0
			indx.AllowReverseScans = True
		Next indx
		Debug.Print " "
	Next ent
End Sub
