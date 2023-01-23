Attribute VB_Name = "basImport"
' Import multiple bas files into current VBE project

Sub Test()
    Dim BasFolder As String
    Dim BasFile As String
    Dim WshShell As Object
    Dim objShell As Object
    Dim objFolder As Object
    Dim MyFile As Object
    
    'BasFolder = "C:\BasFiles"
    BasFolder = "Z:\VSCode-projects\GitHub\SMF_Add-in\code"
    BasFile = "MyMod.bas"
    
    Set WshShell = CreateObject("WScript.Shell")
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace("" & BasFolder & "")
    
    For Each MyFile In objFolder.items
        If MyFile = BasFile Then
            MyQ = MsgBox("Comment : " & objFolder.GetDetailsOf(MyFile, 5) _
                   & vbCrLf & "Do you want to import this module to your project ?", vbYesNo)
            If MyQ = vbYes Then
                ActiveWorkbook.VBProject.VBComponents.Import _
                   BasFolder & Application.PathSeparator & BasFile
            End If
        End If
    Next
    
    Set objFolder = Nothing
    Set objShell = Nothing
    Set WshShell = Nothing
End Sub

Sub ImportPrim()
    Dim Extension, Extensions
    Dim FName As String
       
    'Extensions = Array("cls", "bas", "frm")
    Extensions = Array("bas")
    'BasFolder = "Z:\VSCode-projects\GitHub\SMF_Add-in\code"
    BasFolder = SelectFolder()

    For Each Extension In Extensions
      'FName = Dir(ThisWorkbook.Path & "\*." & Extension)
      FName = Dir(BasFolder & "\*." & Extension)
      Do While FName <> ""
        'ThisWorkbook.VBProject.VBComponents.Import ThisWorkbook.Path & "\" & FName
        ThisWorkbook.VBProject.VBComponents.Import BasFolder & "\" & FName
        FName = Dir
      Loop
    Next
End Sub

Function SelectFolder()
Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        '.InitialFileName = "Z:\VSCode-projects\GitHub\SMF_Add-in\code"
        .InitialFileName = "C:\SMF Add-in"
        .ButtonName = "Select"
        If .show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        SelectFolder = sFolder
    End If
    
End Function

Sub TestFDP()

    foldr = SelectFolder()
    Debug.Print foldr

End Sub
