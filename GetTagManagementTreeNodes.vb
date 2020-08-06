Sub getTreeModel()

Set fso = CreateObject("Scripting.FilesystemObject")

Dim oFile As Object
Dim strLogFile As String

strLogFile = "C:\Temp\GetTagManagementTreeNodesResult.xml"

If fso.FileExists(strLogFile) Then
    'fso.OpenTextFile strLogFile, ForWriting, True
    Set oFile = fso.GetFile(strLogFile)
    Set oFile = oFile.OpenAsTextStream(ForWriting)
Else
    Set oFile = fso.CreateTextFile(strLogFile, False, True)
End If

oFile.WriteLine "<tagStructure>"
    For Each rootItem In NavigationTree.RootNodes
        For Each childItem In rootItem.Childs
            Debug.Print childItem.Name
            oFile.WriteLine "<rootNode>" & childItem.Name
                For Each group In childItem.Childs
                    Debug.Print "   <rootNodeChild>" & group.Name
                    oFile.Write "    <rootNodeChild>" & group.Name
                        For Each connection In group.Childs
                            Debug.Print "       " & connection.Name
                            oFile.WriteLine vbCrLf & "       <connection>" & connection.Name
                            For Each tagGroup In connection.Childs
                                Debug.Print "           " & tagGroup.Name
                                oFile.WriteLine "           <tagGroup grp=""" & tagGroup.Name & """ />"
                            Next
                            oFile.Write "     </connection>"
                        Next
                    oFile.WriteLine "   </rootNodeChild>"
                Next
            oFile.WriteLine "</rootNode>"
        Next
    Next
oFile.WriteLine "</tagStructure>"

oFile.Close
    Debug.Print "yep"

End Sub

