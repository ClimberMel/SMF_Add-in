Attribute VB_Name = "modPictureInsert"
'@Lang VBA
Public Sub RCHInsertPicture(ByVal Target As Range)
    '-----------------------------------------------------------------------------------------------------------*
    ' Subroutine to insert/update a picture within a cell based on cell content ("Image: http://...")
    '-----------------------------------------------------------------------------------------------------------*
    ' 2005.02.01 -- New subroutine; still under consideration/development
    ' 2007.06.26 -- Add "GoTo NextCell" error handling
    '-----------------------------------------------------------------------------------------------> Version 1.2
    ' 2007.06.26 -- Add "GoTo NextCell" error handling
    '-----------------------------------------------------------------------------------------------> Version 2.0a
    ' To automate use:
    '
    '    Private Sub Worksheet_Change(ByVal Target As Range)
    '       Call RCHInsertPicture(Intersect(Target, UsedRange))
    '       End Sub
    '-----------------------------------------------------------------------------------------------------------*
    For Each oCell In Target
        On Error Resume Next
        ActiveSheet.Shapes("Image:" & oCell.Address).Delete
        On Error GoTo NextCell
        If Left(oCell.Value, 7) = "Image: " Then
           With ActiveSheet.Pictures.Insert(Mid(oCell.Value, 8, 999))
                .Left = oCell.Left + 1
                .Top = oCell.Top + 1
                .Name = "Image:" & oCell.Address
                oCell.RowHeight = .Height + 2
                nRatio = oCell.Width / oCell.ColumnWidth
                oCell.ColumnWidth = .Width / nRatio + 2
                End With
           End If
NextCell:
        Next oCell
    End Sub
