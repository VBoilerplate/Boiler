Public Function FolderIsEmpty(myPath As String) As Boolean
    'Checks whether folder is empty    
    FolderIsEmpty = CBool(Dir(myPath & "*.*") = "")
    
End Function
