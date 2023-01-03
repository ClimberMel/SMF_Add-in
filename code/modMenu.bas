Attribute VB_Name = "modMenu"
'-----------------------------------------------------------------------------------------------------------*
' 2014.05.31 -- Originally added all menu routines, thanks go to Andrei Radulescu-Banu
' 2014.06.13 -- Add smfASyncOff and smfASyncOn
' 2017.05.05 -- Add LoadElementsFromFile(21) call
' 2017.05.19 -- Add LoadElementsFromFile(22) call
' 2017.11.08 -- Fix smfMenuRecalculateSelection processing of sWebCache
' 2018.05.02 -- Fix spelling of recalculate
'-----------------------------------------------------------------------------------------------------------*

Private Const sMenuTag As String = "smfCellControlTag"

Private Sub smfAddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl, MySubMenu2 As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call smfDeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup)

    With MySubMenu
        .Caption = "SMF"
        .Tag = sMenuTag
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "smfMenuRecalculateSelection"
            .FaceId = 37
            .Caption = "Recalculate Selection"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "smfForceRecalculation"
            .FaceId = 37
            .Caption = "Recalculate Worksheet"
        End With
        
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "smfFixLinks"
            .FaceId = 5681
            .Caption = "Fix Links"
            .BeginGroup = True
        End With
       
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "smfUpdateDownloadTable"
            .FaceId = 8
            .Caption = "Update Download Table"
        End With
        
        With .Controls.Add(Type:=msoControlButton)
             .OnAction = "smfASyncOn"
             .FaceId = 1664
             .Caption = "Enable Asynchronous processing"
             .BeginGroup = True
             End With
        
        With .Controls.Add(Type:=msoControlButton)
             .OnAction = "smfASyncOff"
             .FaceId = 51
             .Caption = "Disable Asynchronous processing (default)"
             End With
       
        'With .Controls.Add(Type:=msoControlButton)
        '    .OnAction = "smfEnableWebCache"
        '    .FaceId = 1664
        '    .Caption = "Enable Web Cache"
        'End With
       
        'With .Controls.Add(Type:=msoControlButton)
        '    .OnAction = "smfDisableWebCache"
        '    .FaceId = 51
        '    .Caption = "Disable Web Cache"
        'End With
       
        Set MySubMenu2 = .Controls.Add(Type:=msoControlPopup)
        With MySubMenu2
            .Caption = "Logging"
            .Tag = sMenuTag
            .BeginGroup = True
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "smfMenuEnableLog"
                .FaceId = 1664
                .Caption = "Enable"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "smfMenuDisableLog"
                .FaceId = 51
                .Caption = "Disable"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "smfOpenLogFile"
                .FaceId = 1923
                .Caption = "Open Log File"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "smfMenuDeleteLogFile"
                .FaceId = 1668
                .Caption = "Delete Log File"
            End With
        End With
    End With

    ' Add a separator to the Cell context menu.
    'ContextMenu.Controls(4).BeginGroup = True
End Sub

Private Sub smfDeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : "smfCellControlTag".
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = sMenuTag Then
            ctrl.Delete
        End If
    Next ctrl

End Sub

Private Sub smfMenuRecalculateSelection()
    ' Disable the cache and recalculate the selected range
    sWebCache = "N"
    Selection.Dirty
    On Error Resume Next
    Selection.Calculate
    sWebCache = "Y"
End Sub


Private Sub smfMenuEnableLog()
    Call smfLogInternetCalls("Y")
End Sub

Private Sub smfMenuDisableLog()
    Call smfLogInternetCalls("N")
End Sub

Private Sub smfMenuDeleteLogFile()
    Call smfLogInternetCalls("RESET")
End Sub

Sub Auto_Open()
    'Executed when the first workbook is open. This installs the menu.
    Call smfAddToCellMenu
    Call LoadElementsFromFile(21)
    Call LoadElementsFromFile(22)

End Sub

Sub Auto_Close()
    'Executed when the last workbook is closed. This uninstalls the menu
    Call smfDeleteFromCellMenu
End Sub
