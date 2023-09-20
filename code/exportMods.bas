Attribute VB_Name = "exportMods"
Sub xport()

   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to export all modules from the current project.
   ' I have used this to export and save all the project modules from: RCH_Stock_Market_Functions
   '
   ' Reference required to: Microsoft Visual Basic For Applications Extensibility
   '-----------------------------------------------------------------------------------------------------------*
   ' 2023.01.06 -- Created function - Mel Pryor (climbermel@gmail.com)
   '-----------------------------------------------------------------------------------------------------------*
   ' Exports to a defined folder.
   ' Look at having a folder select instead of hard coding it.
   '-----------------------------------------------------------------------------------------------------------*
    
    Dim objMyProj As VBProject
    Dim objVBComp As VBComponent
    
    Set objMyProj = Application.VBE.ActiveVBProject
    
    For Each objVBComp In objMyProj.VBComponents
    If objVBComp.Type = vbext_ct_StdModule Then
    objVBComp.Export "Z:\Temp\vba\test5\" & objVBComp.Name & ".bas"
    End If
    Next

End Sub
