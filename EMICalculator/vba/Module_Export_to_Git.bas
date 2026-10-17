Attribute VB_Name = "Module3"
Option Explicit

Public Sub Export_VBA_To_Git()
    Dim vbComp As Object
    Dim exportPath As String

    ' Change this path to your Git repo folder
    exportPath = ThisWorkbook.Path & "\vba\"

    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    Application.ScreenUpdating = False

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
            Case 2 ' Class Module
                vbComp.Export exportPath & vbComp.Name & ".cls"
            Case 3 ' UserForm
                vbComp.Export exportPath & vbComp.Name & ".frm"
        End Select
    Next vbComp

    Application.ScreenUpdating = False
    MsgBox "VBA Export completed successfully!", vbInformation
End Sub

