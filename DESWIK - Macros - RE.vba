'#Language "WWB.NET"
'Script to join drillholes from the SURVEYEXPORT layer from UGDB into a single polyline for upload to a jigger
'Written by Pat Banks, pat.banks@deswik.com, 11/03/16
'Notes:
' - This version has been written for v2016.1 and will retain RingID and HoleID as vertex attributes, as well as any other vertex attributes
'
'Updates
' - 14/03/16 - rearranged to check active layer visibility BEFORE selecting drillholes

Imports Deswik.Graphics
Imports System.Collections.Generic

Sub Main

    'Set column headers
    Dim headers As New List(Of String)
    headers.Add("RingID")
    headers.Add("HoleID")
    headers.Add("Point 1")
    headers.Add("Point 2")

    'Create a list of attributes to look up data for the table from. Use the first n items in the header list, so updating the script just
    'involves updating one list and one counter
    Dim attLookups As New List(Of String)
    For i As Integer = 0 To 1 'update here if more attributes are added
        attLookups.Add(headers(i))
    Next

    'Ensure the active layer is visible so the output will also be.
    If Not CurrentDoc.MenuCommand.ActiveLayerInvisibleWarning() Then
        Exit Sub
    End If

    'Select polylines to be joined
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select survey export polylines","polyline")
    If sel.Count = 0 Then
        Exit Sub
    End If

    'Create a sorted list of holes
    Dim sortedHoles As list(Of Primaries.Figure) = SortByRingAndHole(sel)

    'Work through the sorted list incrementing the Point Number vertex attribute
    Dim vertCounter As Integer = 1
    For Each fig As Primaries.Figure In sortedHoles
        Dim holePline As Figures.Polyline = fig.asPolyline
        holePline.VerticiesProperties.PropertySet(0, "Point Number", vertCounter)
        holePline.VerticiesProperties.PropertySet(1, "Point Number", vertCounter+1)
        vertCounter +=2
        Next 

    'Fill out the table data
    Dim data As New List(Of List(Of String))
    data.Add(headers)
    For Each holeFig As primaries.Figure In sortedHoles
        Dim row As New List(Of String)
        For Each attName As String In attLookups
            row.Add(holeFig.XProperties.FindName(attName).Value.ToString())
        Next
        'now add point 1 and point 2

        Dim vert1 As String
        Dim vert2 As String

        holeFig.asPolyline.VerticiesProperties.PropertyGet(0,"Point Number",vert1).ToString
        holeFig.asPolyline.VerticiesProperties.PropertyGet(1,"Point Number",vert2).ToString

        row.Add(vert1)
        row.Add(vert2)
        data.Add(row)
    Next

    'select an insertion point
    Dim inspoint As geometry.Point
    If Not CurrentDoc.UserCommands.PointGet(inspoint, CoordinateSystem.TranslatePointConstants.WCS, "Select table insertion point") = commands.User.StatusCode.Success Then
        Exit Sub
    End If

    'Create a table
    Dim tbl As New figures.Table(inspoint, data, CurrentDoc.TextStyles.Default.Height)
    tbl.AlignToView = True 'defaults to false so delete this if that is preferred
    CurrentDoc.ModelSpace.Entities.Add(tbl)
    CurrentDoc.Redraw(True)

    'MsgBox("Done")

End Sub

Function SortByRingAndHole(ByVal sel As Collections.Selection) As List(Of Primaries.Figure)
    Dim result As List(Of Primaries.Figure) = sel.asFiguresList
    Dim sortfields() As String = {"RingID","HoleID"}
    Dim comp As New Deswik.Graphics.Primaries.FigureComparer(sortfields,True)
    result.Sort(comp)
    Return result
End Function




'#Language "WWB.NET"

Imports Deswik.Graphics
Imports Deswik.Graphics.Objects
Imports System.Collections.Generic
Imports System

Sub Main

    Begin Dialog UserDialog 540,98 ' %GRID:10,7,1,1
        Text 30,14,480,21,"Setout Name",.Text2
        OKButton 30,70,250,21
        CancelButton 280,70,230,21
        TextBox 30,42,480,21,.TextBox1
    End Dialog

    Dim dlg As UserDialog
    On Error Resume Next
    Dialog dlg
    If Error=102 Then
    End If

    Dim atv As String = dlg.TextBox1

    CurrentDoc.GlobalConstants.Remove("Setout")

    CurrentDoc.GlobalConstants.Add("Setout",Deswik.Common.Core.Serializable.GlobalConstants.GlobalConstantTypes.String,atv.ToString,"Name of File being created","")

End Sub




FORMAT(DATEADD(BLOCKTEXTVALUE("Date", "dd MM yy"),"3","4"),"dd MMM yyyy")




'Sets overlay type to foreground and aligns to view (for a table)
'#Language "WWB.NET"

Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    'Select The table to bring to the front (Ends Macro if none selected)
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select a Table to Bring to Front.", "table")
    If sel.Count = 0 Then
        Exit Sub
    End If

    'Set "Overlay Type" Property to "FOREGROUND"
    sel.EntitiesPropertiesSet("OverlayType", 2)
    'Set "Align To View" Property to True
    sel.EntitiesPropertiesSet("AlignToView", 1)

End Sub




'Reload Data Source In children
'#Language "WWB.NET"

Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    CurrentDoc.UserCommands.LayerDatasourceReloadChildren()
End Sub




'Changes string ID and description for Floor / Breakline or Active Laser to Inactive.
'Russell Easton 240320
'Changes "1 - Floor" to "101 - old floor"
'Changes "2 - breakline" to "102 - old breakline"
'Changes laser to inactive

'#Language "WWB.NET"

Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    
    'If no polyline is selected, prompt the user to select one.
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Polyline to Change", "polyline")

    'If no line is selected, end the script
    If sel.Count = 0 Then
        Exit Sub
    End If

    'Get the Code of the polyline
    Dim code As Integer
    code = sel.EntitiesAttributesValuesGet("Code").Item(0)

    'Change Attributes based on the code
    Select code
        Case 1 'Floor String
            sel.EntitiesAttributesSet("Code", 101)
            sel.EntitiesAttributesSet("Description", "101 - Old Floor")
        Case 2 'Breakline String
            sel.EntitiesAttributesSet("Code", 102)
            sel.EntitiesAttributesSet("Description", "102 - Old Breakline")
        Case 6 'Laser
            sel.EntitiesAttributesSet("Active", False)
    End Select
End Sub



'#Language "WWB.NET"
'Russell Easton 240321
Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    'Select Polyline and setup variables
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Polyline", "polyline")
    Dim code As String
    Dim code1 As String
    Dim code2 As String
    Dim code2NoZ As String
    Dim code3 As String
    Dim code4 As String
    Dim path As String
    If sel.Count = 0 Then
        Exit Sub
    End If
    'Gets the AMO number from polyline attributes and splits it into separate codes
    code = sel.EntitiesAttributesValuesGet("AMO NUMBER").Item(0)
    code1 = Split(code, "_")(0)
    code2 = Split(code, "_")(1)
    code2NoZ = Replace(code2, "Z", "")
    code3 = Split(code, "_")(2)
    code4 = Split(code, "_")(3)
    'Navigate to folder and open
    Select code1
    Case "AMO"
        'Opens pdf
        path = "U:\04. Mining\04. Mine Planning\09. AMO APPROVED\AMO_"+code2NoZ+"_"+code3+"_"+code4+".pdf"
        Shell("""" & "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" & """ --app=""" & path & """")
    Case "MI"
        'Opens pdf
        path = "U:\04. Mining\04. Mine Planning\10. MIs APPROVED\MI_"+code2NoZ+"_"+code3+"_"+code4+".pdf"
        Shell("""" & "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" & """ --app=""" & path & """")
    Case Else
        'Opens folder
        Select code2
        Case "Z135E"
            path = "U:\04. Mining\04. Mine Planning\08. LEVEL PLANS APPROVED\Z135 EVA\"+code+"\"
        Case "Z135W"
            path = "U:\04. Mining\04. Mine Planning\08. LEVEL PLANS APPROVED\Z135 WVA\"+code+"\"
        Case Else
            path = "U:\04. Mining\04. Mine Planning\08. LEVEL PLANS APPROVED\"+code2+"\"+code+"\"
        End Select
        Shell("explorer.exe """ & path & """", VbAppWinStyle.vbNormalFocus)
    End Select

End Sub









'#Language "WWB.NET"
'Russell Easton 240321
Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    'Select Polyline and setup variables
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Polyline", "polyline")
    Dim code As String
    Dim code1 As String
    Dim code2 As String
    Dim code2NoZ As String
    Dim code3 As String
    Dim code4 As String
    Dim path As String
    If sel.Count = 0 Then
        Exit Sub
    End If
    'Gets the AMO number from polyline attributes and splits it into separate codes
    code = sel.EntitiesAttributesValuesGet("LEVEL").Item(0)
    area = Split(code, "")(3)
    'Navigate to folder and open
    Select area
        Case "0"
            'Opens folder
            path = "K:\06-Master\01 Design\01-Approved Development Design\FTN\"+code+" Level\"
        Case "1"
            path = "K:\06-Master\01 Design\01-Approved Development Design\DRK\"
        Case "2"
            path = "K:\06-Master\01 Design\01-Approved Development Design\GIB\"+code+"\"
        Case "3"
            path = "K:\06-Master\01 Design\01-Approved Development Design\DUN\"+code+" Level\"
        Case "4"
            path = "K:\06-Master\01 Design\01-Approved Development Design\NEL\"+code+"\"
        Case Else

    End Select
    Shell("explorer.exe """ & path & """", VbAppWinStyle.vbNormalFocus)

End Sub







'#Language "WWB.NET"
'Russell Easton 240321
Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    'Select Polyline and setup variables
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Polyline", "polyline")
    Dim code As String
    Dim path As String
    If sel.Count = 0 Then
        Exit Sub
    End If
    'Gets the AMO number from polyline attributes and splits it into separate codes
    code = sel.EntitiesAttributesValuesGet("LEVEL").Item(0)
    Dim area() As Integer

    ReDim area(code.Length - 1)

    For i As Integer = 0 To code.Length - 1
        area(i) = AscW(code(i)) - AscW("0")
    Next i
    Dim a2 As String
    a2 = area(3).ToString()
    'Navigate to folder and open
    Select a2
        Case "0"
            'Opens folder
            path = "K:\06-Master\01 Design\01-Approved Development Design\FTN\"+code+" Level\"
        Case "1"
            path = "K:\06-Master\01 Design\01-Approved Development Design\DRK\"
        Case "2"
            path = "K:\06-Master\01 Design\01-Approved Development Design\GIB\"+code+"\"
        Case "3"
            path = "K:\06-Master\01 Design\01-Approved Development Design\DUN\"+code+" Level\"
        Case "4"
            path = "K:\06-Master\01 Design\01-Approved Development Design\NEL\"+code+"\"
        Case Else
            
    End Select
    Shell("explorer.exe """ & path & """", VbAppWinStyle.vbNormalFocus)

End Sub






'This macro was designed to find any errors within the stope model
'It checks:
'   All Attributes are filled in
'   Panel name and Level are the same
'   Breakthrough level set to N/A for NO breakthrough
'   Breakthrough level not set to N/A for YES breakthrough
'   Stope colour is correct - 33 for PASTE, 0 for VOID
'   Volume is over 100 - This is to check for invalid solids and any small solids that shouldn't exist

'Russell Easton 240330

'#Language "WWB.NET"

Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    
    'Set the selection to all selected polyfaces (solids)
    Dim selection As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Solids to Interrogate", "polyface")
    If selection.Count = 0 Then
        Exit Sub
    End If

    'Initialise variables
    Dim failed As Boolean 'Failure condition
    'Attributes
    Dim breakthrough As String
    Dim breakthroughLevel As String
    Dim dateProcessed As String
    Dim level As String
    Dim mineArea As String
    Dim panelName As String
    Dim stopeStage As String
    Dim volume As Integer
    'For isolating level numbers
    Dim panelDigits As String
    Dim levelDigits As String
    'Counts # of Errors
    Dim count As Integer
    'Error Messages
    Dim errorMessages As New List(Of String)()

    'Loop through all selected stopes
    For Each stope As Primaries.Figure In selection
        
        'Reset failure condition
        failed = False

        'Set attribute variables
        breakthrough = stope.XProperties.FindName("Breakthrough").ToString
        breakthroughLevel = stope.XProperties.FindName("Breakthrough Level").ToString
        dateProcessed = stope.XProperties.FindName("Date Processed").ToString
        level = stope.XProperties.FindName("LEVEL").ToString
        mineArea = stope.XProperties.FindName("Mine Area").ToString
        panelName = stope.XProperties.FindName("Panel Name").ToString
        stopeStage = stope.XProperties.FindName("Stope Stage").ToString
        volume = stope.EntityPropertyValueGet("Volume")


        'Check for blank Attributes
        If breakthrough = "Breakthrough - " Or breakthroughLevel = "Breakthrough Level - " Or dateProcessed = "Date Processed - " Or level = "LEVEL - " Or mineArea = "Mine Area - " Or panelName = "Panel Name - " Or stopeStage = "Stope Stage - "Then
            failed = True
            errorMessages.Add(stope.ToString + " has blank attributes")
        End If

        'Check Panel name and Level are the same
        panelDigits = Left(Right(panelName, Len(panelName) - InStr(panelName, "-") - 1), 4).toString
        levelDigits = Mid(level, InStr(level, "-") + 2, 4).toString
        If Not panelDigits = levelDigits Then
            failed = True
            errorMessages.Add(stope.ToString + " has a 'Panel Name' that does not match the 'LEVEL'")
        End If

        'Check breakthrough levels
        If (breakthrough = "Breakthrough - NO" And Not breakthroughLevel = "Breakthrough Level - N/A") Or (breakthrough = "Breakthrough - YES" And breakthroughLevel = "Breakthrough Level - N/A") Then
            failed = True
            errorMessages.Add(stope.ToString + " has inconsistencies with 'Breakthrough' and 'Breakthrough Level'")
        End If

        'Check pen color - 33 = Paste, 0 = Void
        If (stopeStage = "Stope Stage - PASTE" And Not stope.PenColor.ToString = "33") Or (stopeStage = "Stope Stage - VOID" And Not stope.PenColor.ToString = "0") Then
            failed = True
            errorMessages.Add(stope.ToString + " has the incorrect colour, double check if it is VOID or PASTE")
        End If

        'Check Volume >= 100
        If volume < 100 Then
            failed = True
            errorMessages.Add(stope.ToString + " has a very small volume")
        End If


        'Hide if everything is good with this stope
        If failed = False Then
            stope.Visible = False
        Else
            count += 1
        End If
        

    'Check the next stope
    Next

    'Check # of Errors and lists each error
    If count = 0 Then
        selection.setVisible() 'If no errors, set all to visible
        CurrentDoc.OnMessageOutput("There are No Errors in Stope Attributes", False, False, System.Drawing.Color.Green, True, False)
    Else
        CurrentDoc.OnMessageOutput("There are "+count.ToString+" Errors in Stope Attributes:", False, False, System.Drawing.Color.Red, True, False)
        For Each errorMessage As String In errorMessages
            CurrentDoc.OnMessageOutput(errorMessage, False, False, System.Drawing.Color.Red, True, False)
        Next
    End If

End Sub




'This macro was designed to find any errors within the stope model
'It checks:
'   All Attributes are filled in
'   Panel name and Level are the same
'   Breakthrough level set to N/A for NO breakthrough
'   Breakthrough level not set to N/A for YES breakthrough
'   Stope colour is correct - 33 for PASTE, 0 for VOID
'   Volume is over 100 - This is to check for invalid solids and any small solids that shouldn't exist

'Russell Easton 240330

'#Language "WWB.NET"

Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    
    'Set the selection to all selected polyfaces (solids)
    Dim selection As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Solids to Interrogate", "polyface")
    If selection.Count = 0 Then
        Exit Sub
    End If

    'Initialise variables
    Dim failedError As Boolean 'Failure condition
    Dim failedWarning As Boolean
    'Attributes
    Dim fillType As String
    Dim firingStatus As String
    Dim level As String
    Dim mine As String
    Dim mineArea As String
    Dim minedMonth As String
    Dim minedYear As String
    Dim panelName As String
    Dim scanType As String
    Dim surveyedDate As String
    Dim surveyedBy As String
    Dim volume As Integer
    'For isolating level numbers
    Dim panelDigits As String
    Dim levelDigits As String
    'Counts # of Errors
    Dim errorCount As Integer
    Dim warningCount As Integer
    'Error Messages
    Dim errorMessages As New List(Of String)()
    Dim warningMessages As New List(Of String)()

    'Loop through all selected stopes
    For Each stope As Primaries.Figure In selection
        
        'Reset failure condition
        failedError = False
        failedWarning = False

        'Set attribute variables
        fillType = stope.XProperties.FindName("Fill Type").ToString
        firingStatus = stope.XProperties.FindName("Firing Status").ToString
        level = stope.XProperties.FindName("Level").ToString
        mine = stope.XProperties.FindName("Mine").ToString
        mineArea = stope.XProperties.FindName("Mine Area").ToString
        minedMonth = stope.XProperties.FindName("Mined Month").ToString
        minedYear = stope.XProperties.FindName("Mined Year").ToString
        panelName = stope.XProperties.FindName("Panel Name").ToString
        scanType = stope.XProperties.FindName("Scan Type").ToString
        surveyedDate = stope.XProperties.FindName("Surveyed Date").ToString
        surveyedBy = stope.XProperties.FindName("Surveyed By").ToString
        volume = stope.EntityPropertyValueGet("Volume")


        'Check for blank Attributes
        If fillType = "Fill Type - " Or firingStatus = "Firing Status - " Or level = "Level - " Or level = "LEVEL - " Or mine = "Mine - " Or mineArea = "Mine Area - " Or minedMonth = "Mined Month - " Or minedYear = "Mined Year - " Or panelName = "Panel Name - " Or scanType = "Scan Type - " Or surveyedDate = "Surveyed Date - " Then
            failedError = True
            errorMessages.Add(stope.ToString + " has blank attributes")
        End If

        'Check Panel name and Level are the same
        If Not panelName = "Panel Name - " And Not level = "Level - " Then
            panelDigits = Left(Right(panelName, Len(panelName) - InStr(panelName, "-") - 1), 4).toString
            levelDigits = Mid(level, InStr(level, "-") + 2, 4).toString
            If Not panelDigits = levelDigits Then
                failedWarning = True
                warningMessages.Add(stope.ToString + " has a 'Panel Name' that does not match the 'LEVEL'")
            End If
        End If



        'Check breakthrough levels
        'If (breakthrough = "Breakthrough - NO" And Not breakthroughLevel = "Breakthrough Level - N/A") Or (breakthrough = "Breakthrough - YES" And breakthroughLevel = "Breakthrough Level - N/A") Then
            'failed = True
            'errorMessages.Add(stope.ToString + " has inconsistencies with 'Breakthrough' and 'Breakthrough Level'")
       ' End If

        'Check pen color - 33 = Paste, 0 = Void
        If firingStatus = "Firing Status - FINAL" And Not surveyedBy = "Surveyed By - exception" Then
            If ((fillType = "Fill Type - Paste Filled" And Not stope.PenColor.ToString = "36") Or (fillType = "Fill Type - Open Void" And Not stope.PenColor.ToString = "7") Or (fillType = "Fill Type - Rock Filled" And Not stope.PenColor.ToString = "33")) Then
                failedError = True
                errorMessages.Add(stope.ToString + " has the incorrect colour, double check if it is VOID, PASTE or Rock Filled")
            End If
        End If


        'Check Volume >= 100
        If volume < 100 Then
            failedError = True
            errorMessages.Add(stope.ToString + " has a very small volume")
        End If


        'Hide if everything is good with this stope otherwise add up errors
        If failedError = True Then
            errorCount += 1
        End If
        If failedWarning = True Then
            warningCount += 1
        End If
        If failedError = False And failedWarning = False Then
            stope.Visible = False
        End If

    'Check the next stope
    Next

    'Check # of Errors and lists each error
    If errorCount = 0 And warningCount = 0 Then
        selection.setVisible() 'If no errors, set all to visible
        CurrentDoc.OnMessageOutput("There are No Errors in Stope Attributes", False, False, System.Drawing.Color.Green, True, False)
    Else
        CurrentDoc.OnMessageOutput("There are ", False, False, System.Drawing.Color.Black, True, False)
        CurrentDoc.OnMessageOutput(errorCount.ToString+" Errors ", False, False, System.Drawing.Color.Red, False, False)
        CurrentDoc.OnMessageOutput("and ", False, False, System.Drawing.Color.Black, False, False)
        CurrentDoc.OnMessageOutput(warningCount.ToString+" Warnings ", False, False, System.Drawing.Color.Orange, False, False)
        CurrentDoc.OnMessageOutput("In Stope Attributes:", False, False, System.Drawing.Color.Black, False, False)
        For Each errorMessage As String In errorMessages
            CurrentDoc.OnMessageOutput(errorMessage, False, False, System.Drawing.Color.Red, True, False)
        Next
        For Each warningMessage As String In warningMessages
            CurrentDoc.OnMessageOutput(warningMessage, False, False, System.Drawing.Color.Orange, True, False)
        Next
    End If

End Sub



'Uses the polyline that was just digitised.
'Gets the area and height attributes and multiplies them to get a volume
'Shows the result in the Output window

'Created by Russell Easton 240401

'#Language "WWB.NET"

Imports Deswik.Graphics
Imports System.Collections.Generic
Imports Deswik.Graphics.DataDrivers
Imports System.Collections

Sub Main
    
    'Set up Variables
    Dim area As Double
    Dim height As Double
    Dim volume As Double

    'Gets the selected polyline
    Dim sel As Collections.Selection = CurrentDoc.UserCommands.SelectionInteractiveGet("Select Polyline", "polyline")
    If sel.Count = 0 Then
        Exit Sub
    End If

    For Each strip As Primaries.Figure In sel
        
        'sets the variables
        On Error GoTo ErrorHandler
            area = strip.EntityPropertyValueGet("Area")
            height = sel.EntitiesAttributesValuesGet("Strip Height").Item(0)
            volume = Round(area * height, 3)

            'Write a message In the Output window showing volume
            CurrentDoc.OnMessageOutput("Strip Volume: "+(volume).ToString+"m³", False, False, System.Drawing.Color.Green, True, False)

            'Move to Stripping Layer
            CurrentDoc.UserCommands.EntitiesCopyMoveToLayer(False, sel, {"ASBUILT\EOM\STRIPPING"}, True, False, False)

        Exit Sub
        ErrorHandler:
            CurrentDoc.OnMessageOutput("Error: Check height is entered", False, False, System.Drawing.Color.Red, True, False)

    Next


End Sub






