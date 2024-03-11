Attribute VB_Name = "exportMods"
Sub xport()

    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to export all modules from the current project.
    ' I have used this to export and save all the project modules from: RCH_Stock_Market_Functions
    '
    ' Reference required to: Microsoft Visual Basic For Applications Extensibility
    '-----------------------------------------------------------------------------------------------------------*
    ' 2023.01.06 -- Created function - Mel Pryor (climbermel@gmail.com)
    '
    ' 2024.03.11 -- added code to create a dated folder to export code to instead of hard coding it as I had
    '               this way I don't forget and overwrite the existing code
    '-----------------------------------------------------------------------------------------------------------*
    ' -- Exports to a dated folder.
    ' -- Could have it check if folder exists so it doesn't error if I rerun it the same day for some reason
    ' -- Could change to \\NASDCC042\Software\Trading\SMF Add-in\Code Archive
    '    if I adjust folder naming...
    '-----------------------------------------------------------------------------------------------------------*
    
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    Dim dtToday As String
    
    dtToday = Format(Date, "yyyy-mm-dd")
    Folder = "Z:\Temp\vba\" & dtToday
    
    foldernm = Folder & "\"
    Set objMyProj = Application.VBE.ActiveVBProject
    
    If fso.FolderExists(folderPath) Then
        MkDir Folder    
        For Each objVBComp In objMyProj.VBComponents
            If objVBComp.Type = vbext_ct_StdModule Then
                objVBComp.Export foldernm & objVBComp.Name & ".bas"
            End If
        Next
    Else
        MsgBox "Folder already exists!"
    End If

End Sub
