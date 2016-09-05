'TITLE:  Entity Defintions
'DESCRIPTION:  This macro will list all the tables /Entities a
' that have spaces in name
'AUTHOR:  Karen Lopez Based on macro by Jason Tiret
'EMAIL: karen@infoadvisors.com
'DATE:  11/1/2008
'UPDATED:

Option Explicit



'ERStudio global variables
Dim MyModel As Model
Dim MyDiagram As Diagram
Dim MyEntity As Entity
Dim MyAttribute As AttributeObj
Dim EntCount As Integer
Dim EntProbCount As Integer
Dim TNLength As Integer





Sub Main
	'initialize ER variables
	Set MyDiagram = DiagramManager.ActiveDiagram
	Set MyModel = MyDiagram.ActiveModel


	EntCount = 0
	EntProbCount = 0


    Debug.Clear


	'loop through each entity in the active model
	For Each MyEntity In MyModel.Entities


			If InStr(1,MyEntity.TableName, " ") > 0 Then

				Debug.Print  MyEntity.TableName

			EntProbCount = EntProbCount +1
			End If



	EntCount = EntCount + 1


	Next MyEntity

Debug.Print EntProbCount & " Entities found out of " & EntCount & " Total Entities with problems"
	Wait (100)
End Sub
