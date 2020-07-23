Sub MoveRClickDelKeyboard()

Rem *******************************************************************************************************
Rem This program cycles through all .pdl files in your folder of choice. For each page it opens it will cycle through all objects it finds in the page.
Rem The program checks for actions on each mouse or keyboard event.
Rem If there is an action on the right click event, the action is moved to the left click event if the left action doesn't exist and if
Rem the object calls a faceplace upon right mouse click event.
Rem Subsequently, the keyboard actions are deleted.
Rem Each action has a control variable with which the program can see if the object currently processed has been changed at all.
Rem If the page has had any object changed, the .pdl file will be saved.
Rem The end of the script runtime will yield a message showing how many seconds the action took and how many objects have been processed.
Rem Checking whether the .pdl files were saved is a good indicator of whether the files had any objects changed.
Rem
Rem ensure Microsoft Scripting Runtime is checked in Tools\References menu
Rem
Rem             |------------|------------|------------|
Rem             |   Author   |    Date    |   Version  |
Rem             |   MURA02   | 08.01.2019 |     1.0    |
Rem             |------------|------------|------------|
Rem
Rem *******************************************************************************************************

Dim objDocument As Document
Dim objVBScript As HMIScriptInfo
Dim FileName, Path As String
Dim FolderPath, strSourceCode As String
Dim fso As FileSystemObject
Dim fdr As Scripting.Folder
Dim f As File
Dim fs As Files
Dim param, objsChanged As Integer
Dim propexists As Boolean
Dim hasProperty, hasFaceplateD, hasFaceplateU As Integer
Dim hasMouseEvent, hasLeftDownMouseEvent, hasRightDownMouseEvent, hasLeftUpMouseEvent, hasRightUpMouseEvent, hasKeyPressEvent, hasKeyRlsEvent As Integer
Dim filterPages(10)
Dim z


StartTime = Timer()
 
Rem Application.VBE.ActiveVBProject.References.AddFromFile "C:\Windows\Syswow64\scrrun.dll"

Set fso = CreateObject("Scripting.FilesystemObject")

Path = "C:\Project\Processed"
FolderPath = Path

Set fdr = fso.GetFolder(FolderPath)
Set fs = fdr.Files

Rem initialize objects processed counter
objsChanged = 0

Rem define the array of multiple files of choice
filterPages(0) = "HFM#27-02-01_n_#HFM-DevicesTestCC-2"
filterPages(1) = "HFM#69-02-00_w_#MED-EmulsionSystemFM.pdl"
filterPages(2) = "CRM#64-03-00_w_#MED-FilterPumps.pdl"
filterPages(3) = "CRM#61-01-00_f_#MED-HighPressureHydraulic.pdl"
filterPages(4) = ""
filterPages(5) = ""
filterPages(6) = ""
filterPages(7) = ""
filterPages(8) = ""
filterPages(9) = ""


Rem default filters
filterpdl = ".pdl"
filtere = "_e_#"
filterf = "_f_#"
filterp = "_p_#"
filtern = "_n_#"
filterw = "_w_#"
filtersys1 = "@"
filtersys2 = "i8Sys"
filtersys3 = "i8Int"
filtersys4 = "#sys"
filterapp = "i8App"
filteri8 = "i8"
filterObjectType = "HMIPictureWindow"
filterObjectName = "@V3_SMS_StdDev_Pmp_Up_Fph_2"


For Each f In fs

Rem init the property change counter to have a indicator for saving
k = 0

Rem Filter Files
Rem rem multiple page filter *************************************************************
'z = 0
'For z = 0 To (UBound(filterPages) - 1)
'If filterPages(z) <> "" And InStr(1, f, filterPages(z), vbTextCompare) > 0 Then 'filter to single out multiple pages, comment if not needed
Rem **********************************************************************************
If InStr(1, f, filterpdl, vbTextCompare) > 0 Then  'filter through .pdl files only
'If InStr(1, f, filtern, vbTextCompare) > 0 Or InStr(1, f, filterw, vbTextCompare) > 0 Or InStr(1, f, filterp, vbTextCompare) > 0 Or InStr(1, f, filtere, vbTextCompare) > 0 Or InStr(1, f, filterf, vbTextCompare) > 0 Then     'filter through types of pages
'If InStr(1, f, filtersys1, vbTextCompare) = 0 Or InStr(1, f, filtersys2, vbTextCompare) = 0 Or InStr(1, f, filtersys3, vbTextCompare) = 0 Or InStr(1, f, filtersys4, vbTextCompare) = 0 Or InStr(1, f, filteri8, vbTextCompare) = 0 Then       'filter out system pages et. al.
Rem end of file filter

    Application.Documents.Open f, hmiOpenDocumentTypeVisible
    yyyy = f.Name
    
    Dim obj As HMIObject
    Dim objs As HMIObjects
    
    Set objs = ActiveDocument.HMIObjects
        For Each obj In objs 'cycle through objects to check and perform changes
        
'        If obj.ObjectName = filterObjectName Then ' uncomment if need be to test for one object
        
        Rem script to be performed on the filtered pages
            
            Rem initialize checkbits
            hasProperty = 0
            hasMouseEvent = 0
            hasLeftDownMouseEvent = 0
            hasRightDownMouseEvent = 0
            hasLeftUpMouseEvent = 0
            hasRightUpMouseEvent = 0
            hasKeyPressEvent = 0
            hasKeyRlsEvent = 0
            hasFaceplate = 0
                        
            Rem check if these objects have actions on these events (any type)
            For j = 1 To obj.Events.Count
                If obj.Events(j).EventName = "OnRButtonDown" Then
                    If obj.Events.Item("OnRButtonDown").Actions.Count > 0 Then
                        hasRightDownMouseEvent = hasRightDownMouseEvent + 1
                    End If
                End If
                If obj.Events(j).EventName = "OnRButtonUp" Then
                    If obj.Events.Item("OnRButtonUp").Actions.Count > 0 Then
                        hasRightUpMouseEvent = hasRightUpMouseEvent + 1
                    End If
                End If
                If obj.Events(j).EventName = "OnLButtonDown" Then
                    If obj.Events.Item("OnLButtonDown").Actions.Count > 0 Then
                        hasLeftDownMouseEvent = hasLeftDownMouseEvent + 1
                    End If
                End If
                If obj.Events(j).EventName = "OnLButtonUp" Then
                    If obj.Events.Item("OnLButtonUp").Actions.Count > 0 Then
                        hasLeftUpMouseEvent = hasLeftUpMouseEvent + 1
                    End If
                End If
                If obj.Events(j).EventName = "OnKeyUp" Then
                    If obj.Events.Item("OnKeyUp").Actions.Count > 0 Then
                        hasKeyRlsEvent = hasKeyRlsEvent + 1
                    End If
                End If
                If obj.Events(j).EventName = "OnKeyDown" Then
                    If obj.Events.Item("OnKeyDown").Actions.Count > 0 Then
                        hasKeyPressEvent = hasKeyPressEvent + 1
                    End If
                End If
            Next
            
            If hasRightDownMouseEvent > 0 And hasLeftDownMouseEvent = 0 Then
            On Error Resume Next
                Rem get right down mouse action
                hasFaceplate = 0
                Set objVBScript = obj.Events.Item("OnRButtonDown").Actions.Item(1)
                    With objVBScript
                        strSourceCode = .sourceCode
                        Rem check if object calls faceplate
                        If InStr(1, strSourceCode, "rtFpCallUp", vbTextCompare) Then 'check if object calls a faceplate on right click press
                            hasFaceplateD = hasFaceplateD + 1
                        End If
                    End With
                Rem inject right down mouse action to left down mouse action
                If hasFaceplateD > 0 Then
                obj.Events.Item("OnRButtonDown").Actions.Item(1).Delete 'DELETE RIGHT CLICK PRESS
                Set objVBScript = obj.Events.Item("OnLButtonDown").Actions.AddAction(hmiActionCreationTypeVBScript)
                    With objVBScript
                        .sourceCode = strSourceCode
                    objsChanged = objsChanged + 1
                    k = k + 1 'increment indicator to see if any change has been performed for saving
                    End With
                End If
            End If
            
            If hasFaceplateD > 0 Then
                obj.Events.Item("OnRButtonDown").Actions.Item(1).Delete 'DELETE RIGHT CLICK PRESS
            End If
            
            If hasRightUpMouseEvent > 0 And hasLeftUpMouseEvent = 0 Then
            On Error Resume Next
                Rem get right up mouse action
                hasFaceplate = 0
                Set objVBScript = obj.Events.Item("OnRButtonUp").Actions.Item(1)
                    With objVBScript
                        strSourceCode = .sourceCode
                        Rem check if object calls faceplate
                        If InStr(1, strSourceCode, "rtFpCallUp", vbTextCompare) Then 'check if object calls a faceplate on right click release
                            hasFaceplateU = hasFaceplateU + 1
                        End If
                    End With
                Rem inject right up mouse action to left down mouse action
                If hasFaceplateU > 0 Then
                Set objVBScript = obj.Events.Item("OnLButtonUp").Actions.AddAction(hmiActionCreationTypeVBScript)
                    With objVBScript
                        .sourceCode = strSourceCode
                    objsChanged = objsChanged + 1
                    k = k + 1 'increment indicator to see if any change has been performed for saving
                    End With
                End If
            End If
            If hasFaceplateU > 0 Then
                obj.Events.Item("OnRButtonUp").Actions.Item(1).Delete 'DELETE RIGHT CLICK RELEASE
            End If
            
            If hasKeyPressEvent > 0 Then
                On Error Resume Next
                obj.Events.Item("OnKeyDown").Actions.Item(1).Delete 'DELETE KEYBOARD PRESS
                If hasKeyPressEvent > 0 Then
                    keysDeleted = keysDeleted
                Else
                    keysDeleted = keysDeleted + 1
                End If
                    k = k + 1 'increment indicator to see if any change has been performed for saving
            End If
            
            If hasKeyRlsEvent > 0 Then
                On Error Resume Next
                obj.Events.Item("OnKeyUp").Actions.Item(1).Delete 'DELETE KEYBOARD RELEASE
                If hasKeyRlsEvent > 0 Then
                    keysDeleted = keysDeleted
                Else
                    keysDeleted = keysDeleted + 1
                End If
                    k = k + 1 'increment indicator to see if any change has been performed for saving
            End If
            
        Rem end of script to be performed
        
'        End If 'end object filter
        
        Next obj
    
    Rem check if the file needs to be saved or not, else it will just close it
    If k > 0 Then
'        MsgBox "modified"
        ActiveDocument.Save
    End If
    Set objs = Nothing
    ActiveDocument.Close
    
'End If
'End If
End If
Rem end multiple page filter *************************************************************
'End If 'end multiple file filter
'Next   'next row in array of files
Rem ***************************************************************************************


Next 'next file in graCS

Set fso = Nothing
Set Files = Nothing

EndTime = Timer()
MsgBox objsChanged & " objects have had clicks moved " & vbCrLf _
& keysDeleted & " objects have had keyboard actions deleted " & " in " & FormatNumber(EndTime - StartTime, 2) & " seconds"

End Sub
