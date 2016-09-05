
'Author       : Karen Lopez InfoAdvisors.com
'Macro Purpose: Export an image of all submodels in an active model.  Ensure you have selected
'               the proper model before running this macro
'				*** Warning: This macro temporarily removes back ground colours, then reapplies them
'				If you stop the macro mid-run, your background colours could be lost.

Sub Main
	Dim theDiagram As Diagram
	Dim theModel As Model
	Dim theSubModel As SubModel
    Dim sReportName As String
    Dim FileDt As String
    Dim FullDate As Date
    Dim ModelRoot As String
    Dim PathName As String
    Dim LogPhys As String
	Dim HoldBackGround As Long
	Dim FileThere As Boolean
	Dim Vers As String



	Debug.Clear

	Set theDiagram = DiagramManager.ActiveDiagram

	Set theModel = theDiagram.ActiveModel

	FullDate = Now()
	'Can be used to append file date and time to image names.
	'Currently not used in this macro, but retained here for future use.

	FileDt = Trim(Str(DatePart("yyyy",FullDate))) + Trim(Format(FullDate,"mmm")) + Trim(Str(DatePart("d",FullDate))) + "-" + Format(FullDate, "hhmmss")

	'The Diagram Property Version Number will be used in file names
	'Ensure this is set for your model to get the right version number on diagram names.

	Vers = theDiagram.Version
	Vers = Replace(Trim(Vers), ".","_")

	If theModel.Logical = True Then
		LogPhys = "Log-"
	Else
		LogPhys = "Phys-"
	End If
	
	'Set  path to where you want the image files saved
	'Should be customized to your own environment
	PathName = "C:\ERSoutput\"



	ModelRoot = Left(theDiagram.FileName,(InStr(1,theDiagram.FileName,"."))-1)

	For Each theSubModel In theModel.SubModels
		If theSubModel.MainSubModel = False Then

			'Save the Current Background
			HoldBackGround = theSubModel.DiagramBackgroundColor


			theSubModel.DiagramBackgroundColor = 16777215 ' white


			'sReportName = theDiagram.GetFullFilePath
			'PathName = Left(sReportName,(InStr(1,sReportName, theDiagram.FileName)-1))


			sReportName = PathName + ModelRoot + "-" + LogPhys+Replace(Trim(theSubModel.Name)," ","")+ " " + Vers  + ".png"

			'MsgBox(sReportName)
			'sReportName = ModelRoot + "-"+ LogPhys+Replace(Trim(theSubModel.Name)," ","")+"-" + FileDt + ".png"

			'Set Image properties
			theSubModel.ExportImageOptions.Size=1
			theSubModel.ExportImageOptions.ImageType = 4
			theSubModel.ExportImageOptions.ImageQuality = 100
			theSubModel.ExportImageOptions.PreserveAspectRatio = True

			theSubModel.ExportImageOptions.FileName = sReportName
			Debug.Print sReportName

			theSubModel.ExportImageOptions.ExportImage

			Wait 1
			'Reapply real background
			theSubModel.DiagramBackgroundColor = HoldBackGround
		End If
	Next theSubModel
	MsgBox("Model Images Exported")
End Sub
