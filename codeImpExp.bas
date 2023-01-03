Attribute VB_Name = "codeImpExp"
'This will export all modules from the current project.
'I used this to save all the project: RCH_Stock_Market_Functions

Sub x()

' reference to extensibility library

Dim objMyProj As VBProject
Dim objVBComp As VBComponent

Set objMyProj = Application.VBE.ActiveVBProject

For Each objVBComp In objMyProj.VBComponents
If objVBComp.Type = vbext_ct_StdModule Then
objVBComp.Export "Z:\Temp\vba\test2\" & objVBComp.Name & ".bas"
End If
Next

End Sub