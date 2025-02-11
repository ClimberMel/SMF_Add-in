Attribute VB_Name = "modPictureToggle"
'@Lang VBA
Public Sub RCHTogglePicture()
    '-----------------------------------------------------------------------------------------------------------*
    ' Subroutine to toggle image (i.e. zoom in/zoom out) within a cell
    '-----------------------------------------------------------------------------------------------------------*
    ' 2005.02.01 -- New subroutine; still under consideration/development
    '-----------------------------------------------------------------------------------------------> Version 1.2
    ' 2006.12.10 -- Add zOrder option to bring chart to front when displayed
    '-----------------------------------------------------------------------------------------------> Version 1.3
    '-----------------------------------------------------------------------------------------------------------*
    ' First use requires subroutine be executed while a cell with URL of picture is selected.  After that,
    ' clicking on the image zooms it to normal size or back down to normal cell size.
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo NoShape
    Set oShape = ActiveSheet.Shapes(Application.Caller)
    On Error GoTo 0
    sOldURL = oShape.AlternativeText
    sNewURL = oShape.TopLeftCell.Text
    If sOldURL = sNewURL Then
        With oShape
             If Abs(.Height - .TopLeftCell.Height) < 1 Then
                .ScaleHeight 1, msoTrue
                .ScaleWidth 1, msoTrue
             Else
                .Height = .TopLeftCell.Height
                End If
             End With
    Else
       iLeft = oShape.Left
       iTop = oShape.Top
       oShape.Delete
       Set oShape = ActiveSheet.Pictures.Insert(sNewURL)
       oShape.Name = Application.Caller
       oShape.OnAction = "RCHTogglePicture"
       oShape.Left = iLeft
       oShape.Top = iTop
       ActiveSheet.Shapes(Application.Caller).AlternativeText = sNewURL
       End If
    On Error Resume Next
    oShape.ZOrder msoBringToFront
    On Error GoTo 0
    Exit Sub
NoShape:
    sURL = Selection.Text
    Set oShape = ActiveSheet.Pictures.Insert(sURL)
    oShape.OnAction = "RCHTogglePicture"
    oShape.Left = Selection.Left
    oShape.Top = Selection.Top
    ActiveSheet.Shapes(oShape.Name).AlternativeText = sURL
    End Sub
