Attribute VB_Name = "FolderOpenerModule"
Public Sub open_project_folder(fldPath As String)


    ' -----------------------------------------------------------------
    
    
    
    If Trim(fldPath) <> "" Then
    
        Shell "C:\WINDOWS\explorer.exe """ & CStr(Trim(fldPath)) & "", vbNormalFocus
    Else
        MsgBox "nie ma czego otworzyc!"
    End If
    
    ' -----------------------------------------------------------------
End Sub

