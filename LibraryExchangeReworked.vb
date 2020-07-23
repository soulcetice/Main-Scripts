Option Explicit
Sub run()

Call LibraryExchange

End Sub
Function CreateOrOpenLogFile(logFileName As String)

    Dim fso As FileSystemObject
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FilesystemObject")
    
    If fso.FileExists(logFileName) Then
        Set file = fso.GetFile(logFileName)
        Set file = file.OpenAsTextStream(ForWriting)
    Else
        Set file = fso.CreateTextFile(logFileName, False, True)
    End If
    
End Function
Function filterPages()

    Dim filterPages(10)
    
    Rem define the array of multiple files of choice
    filterPages(0) = "HFM#01-01-91_e_#HFM-TrackingOverviewClass3.pdl"
    filterPages(1) = ""
    filterPages(2) = ""
    filterPages(3) = ""
    filterPages(4) = ""
    filterPages(5) = ""
    filterPages(6) = ""
    filterPages(7) = ""
    filterPages(8) = ""
    filterPages(9) = ""
                
        Rem Rem multiple page filter *************************************************************
        'Dim z
        'z = 0
        'For z = 0 To (UBound(filterPages) - 1)
        'If filterPages(z) <> "" And InStr(1, f, filterPages(z), vbbinarycompare) > 0 Then 'filter to single out multiple pages, comment if not needed
        Rem **********************************************************************************
        
End Function
Function renameObjListToStandard(objects As HMIObjects)

Dim namesReplaced As Integer
Dim obj As HMIObject
Dim objName As String

namesReplaced = 0

For Each obj In objects
    If InStr(1, objName, "@V3_SMS_@V3_SMS_", vbBinaryCompare) Then
        objName = Replace(objName, "@V3_SMS_@V3_SMS_", "@V3_SMS_", , , vbBinaryCompare)
    End If
Next

Rem **********************************************************************************************************
Rem rename to new standard
        For Each obj In objects
        If InStr(1, objName, "@SMSMode", vbBinaryCompare) Then
            objName = Replace(objName, "@SMSMode", "@V3_SMS_Mode_", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        If InStr(1, objName, "@SMSSigLmp", vbBinaryCompare) Then
            objName = Replace(objName, "@SMSSigLmp", "@V3_SMS_SignLmp_", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        If InStr(1, objName, "Pb_", vbBinaryCompare) Then
            objName = Replace(objName, "Pb_", "@V3_SMS_Pb_", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        If InStr(1, objName, "@SMS_RefVal", vbBinaryCompare) Then
            objName = Replace(objName, "@SMS_RefVal", "@V3_SMS_DatCls3_RefVal", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        If InStr(1, objName, "@SMS_ActValChar", vbBinaryCompare) Then
            objName = Replace(objName, "@SMS_ActValChar", "@V3_SMS_DatCls3_ActValChar", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        If InStr(1, objName, "@SMS_ActVal", vbBinaryCompare) Then
            objName = Replace(objName, "@SMS_ActVal", "@V3_SMS_DatCls3_ActVal", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        If InStr(1, objName, "@SMSMovement", vbBinaryCompare) Then
            objName = Replace(objName, "@SMSMovement", "@V3_SMS_Movement", , , vbBinaryCompare)
            namesReplaced = namesReplaced + 1
        End If
        Next
Rem end rename to new standard
Rem **********************************************************************************************************
            
End Function

Sub LibraryExchange()
    
    
    Rem *******************************************************************************************************
    Rem This program cycles through all .pdl files in your folder of choice. For each page it opens it will cycle through all objects it finds in the page.
    Rem ensure Microsoft Scripting Runtime is checked in Tools\References menu
    Rem
    Rem             |------------|------------|------------|
    Rem             |   Author   |    Date    |   Version  |
    Rem             |   MURA02   | 30.03.2020 |     2.1    |
    Rem             |------------|------------|------------|
    Rem
    Rem *******************************************************************************************************
    
    On Error Resume Next
    Application.VBE.ActiveVBProject.References.AddFromFile "C:\Windows\SysWOW64\scrrun.dll" 'Microsoft Scripting Runtime dll
    
    Dim objDocument As Document
    Dim FileName, Path, PathLib As String
    Dim FolderPath, FolderPathLib, strSourceCode As String
    Dim fso As FileSystemObject
    Dim fdr As Scripting.Folder
    Dim fdrLib As Scripting.Folder
    Dim f As file
    Dim fLib As file
    Dim fs As Files
    Dim fsLib As Files
    Dim objsChanged As Integer
    Dim objsCreated As Long
    Dim libObjScanned As Long
    Dim objsScanned As Long
    Dim tooltipAdded As Integer
    Dim startTime, endTime
    Dim filterObjectName, filterpdl, filtere, filterf, filterp, filtern, filterw, filtersys1, filtersys2, filtersys3, filtersys4, filterapp, filteri8, filterObjectType As String
    Dim Files As Files
    Dim strLogFile As String
    Dim fdate
    Dim ndate
    
    startTime = Timer()
    
    strLogFile = "C:\Project\HMI-ToConvert-Files\Project\libRefresh_logfile" & startTime & ".txt"
    FolderPath = "C:\Project\HMI-ToConvert-Files\Project"
    FolderPathLib = "C:\Project\ProcessedLib"
    
    Dim oFile As Object
    Set fso = CreateObject("Scripting.FilesystemObject")
    If fso.FileExists(strLogFile) Then
        Set oFile = fso.GetFile(strLogFile)
        Set oFile = oFile.OpenAsTextStream(ForWriting)
    Else
        Set oFile = fso.CreateTextFile(strLogFile, False, True)
    End If
    
    oFile.WriteLine "This file contains the object names for the objects that weren't switched with their library counterparts and their respective .pdl filenames." & vbCrLf
    oFile.WriteLine Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " " & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & vbCrLf
        
    Set fdr = fso.GetFolder(FolderPath)
    Set fdrLib = fso.GetFolder(FolderPathLib)
    Set fs = fdr.Files
    Set fsLib = fdrLib.Files
    
    Rem initialize objects processed counter
    objsChanged = 0
    objsCreated = 0
    libObjScanned = 0
    objsScanned = 0
    tooltipAdded = 0
    
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
        
        If InStr(1, f, filterpdl, vbTextCompare) > 0 Then  'filter through .pdl files only
            
        Rem ** check if date is earlier than * for use to see if file was not recently modified ******
        fdate = CDate(f.DateLastModified)
        ndate = CDate("13/05/2020 10:00")
        If (fdate < ndate) Then
          'MsgBox fdate
        Else
          GoTo nextFile
        End If
        Rem end ** check if date is earlier than * for use to see if file was not recently modified **
        
            On Error Resume Next
            Application.Documents.CloseAll
            Application.Documents.Open f, hmiOpenDocumentTypeVisible
        
            Rem script to be performed on the filtered pages
            
            Rem init the property change counter to have a indicator for saving
            Dim k As Integer
            Dim obj As HMIObject
            Dim objs As HMIObjects
            Dim objLib As HMIObject
            Dim objsLib As HMIObjects
            Dim objNew As HMIObject
            Dim dynamicDialog As HMIDynamicDialog
            Dim dynamicDialog2 As HMIDynamicDialog
            Dim dynamicDialogDest As HMIDynamicDialog
            Dim objVarTrigger As HMIVariableTrigger
            Dim objVBScript As HMIScriptInfo, objVBScriptDest As HMIScriptInfo
            Dim myTrigger
            Dim sameTrig, trigVarName, resType, boolPosVal, boolNegVal, trigType, trigCycleType, trigCycleName, trigCycleTime
            Dim trigNo As Integer
            Dim objGroup As HMIGroup
            Dim objGroups As HMIGroupedObjects
            Dim trigCnt, tempTrig, act, n, m, i, j, a, l, e, anElseCase, triggers
            Dim anRangeTo, anValue As Variant
            Dim sourceCode, dynSourceCode As String
            Dim propertiesArrayMsg, propArr, selectedProperties As String
            Dim dynStateType As HMIDynamicStateType
            Dim varStateType As HMIVariableStateType
            Dim analogCount, trigCount As Integer
            Dim propArrBuild As String
            Dim x As Integer
            Dim trigVarNameArr(10)
            Dim trigCycleTypeArr(10)
            Dim trigName As String
            Dim objLibString As String
            Dim objEvent As HMIEvent
            Dim temp, unitAd, unitSFactor, unitSFormat, unitSOffset As String
            Dim tagChecker, tagCount, checkBit, ctrArr, idxLibArr As Integer
            Dim objLibArr(100)
            Dim objWasChk
            Dim objName As String
            Dim tagUChg
            Dim objTimer
            Dim objEndTimer
            Dim objTime
            
            Set objs = Application.Documents.Item(f).HMIObjects
            
            Call RenameByDatCls(objs)
            
            For Each obj In objs
                
                If obj.ObjectName = "<Method 'ObjectName' of object 'IHMIPolygon' failed>" Or IsNumeric(CInt(obj.Left)) = False Then
                  'MsgBox "empty obj name"
                  oFile.WriteLine f.Name & ", " & objName & ", this object was not exchanged because of failure to access properties Left/ObjectName"
                  GoTo nextObject
                End If
                
                objTimer = Timer()
                
                objName = obj.ObjectName
                
                objsScanned = objsScanned + 1
                
                Dim tagList, tagAdAt, tagAct, tagClAt, tagFrcRls, tagText, tagPnAt, tagPpAt, tagPV, tagLMN, tagS1At, tagS2At, tagS3At, tagS4At, tagActV, tagSetV, tagOutV, tagMnAt, tagMsAt, tagOpMask, tagActPos, tagUnitFactor, tagUnitFormat, tagUnitOffset
                Dim LibTagList, libTagAdAt, libTagActV, libTagSetV, libTagOutV, libTagMnAt, libTagMsAt, libTagOpMask, libTagActPos, libTagClAt, libTagFrcRls, libTagPnAt, libTagPpAt, libTagPV, libTagLMN, libTagS1At, libTagS2At, libTagS3At, libTagS4At, libTagUnitFactor, libTagUnitFormat, libTagUnitOffset, libTagText
                Dim sTagAdAt, sTagClAt, sTagFrcRls, sTagText, sTagPnAt, sTagPpAt, sTagPV, sTagLMN, sTagS1At, sTagS2At, sTagS3At, sTagS4At, sTagActV, sTagSetV, sTagOutV, sTagMnAt, sTagMsAt, sTagOpMask, sTagActPos, sTagUnitFactor, sTagUnitFormat, sTagUnitOffset
                Dim tagStatus 'the newly created status tag
                Dim tagCalculateFb, tagStatusFb
                Dim tagMlcCtrl, tagMlcLevel
                Dim tagTabNo, tagMin, tagMax
                Dim tagRef
                Dim tagHsa
                Dim tagHmo
                Dim tagMlc
                Dim tagRam
                Dim tagFlow
                Dim tagLength
                Dim tagPosition
                Dim tagCounter
                Dim ActVinText, SetVinText As Integer
                Dim tagMtrPlate
                Dim objClAt As HMIObject
                Dim tagListArray() As String
                
Rem *********************************************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
Rem ** LIBRARY DEFINITIONS **********************************************************************************************************************************************************************

            If (InStr(1, objName, "AnaMsr", vbBinaryCompare) Or InStr(1, objName, "DigMsr", vbBinaryCompare) Or InStr(1, objName, "PidOutp", vbBinaryCompare) Or _
            InStr(1, objName, "SenGen", vbBinaryCompare) Or InStr(1, objName, "DigMsr", vbBinaryCompare) Or InStr(1, objName, "PidOutp", vbBinaryCompare) Or _
            InStr(1, objName, "SignLmp_", vbBinaryCompare) Or InStr(1, objName, "PidPV", vbBinaryCompare)) > 0 And InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then
            Set fLib = fsLib.Item("Libi8_n_Devices-Sensors_Tooltip+UAL.pdl") 'for act val or ref val objects, with or w/o main attribute
            Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
              temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
            ElseIf (InStr(1, objName, "SignLmpFault", vbBinaryCompare) Or InStr(1, objName, "SignLmpStatus", vbBinaryCompare) Or InStr(1, objName, "SignLmpWarn", vbBinaryCompare)) > 0 And InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then
                    GoTo nextObject
                    Set fLib = fsLib.Item("Libi8_n_StatusLamps-Binary_Tooltip.pdl") 'for act val or ref val objects, with or w/o main attribute
                    Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                    temp = objName
                    If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                        If IsNumeric(Right(temp, 3)) Then
                            temp = Left(temp, Len(temp) - 3)
                        ElseIf IsNumeric(Right(temp, 2)) Then
                            temp = Left(temp, Len(temp) - 2)
                        ElseIf IsNumeric(Right(temp, 1)) Then
                            temp = Left(temp, Len(temp) - 1)
                        End If
                    Else
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                    End If
            ElseIf (InStr(1, objName, "DatCls2_ActVal", vbBinaryCompare) Or InStr(1, objName, "DatCls2_RefVal", vbBinaryCompare) Or InStr(1, objName, "DatCls3_ActVal", vbBinaryCompare) Or InStr(1, objName, "DatCls3_RefVal", vbBinaryCompare)) > 0 And InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then
                If InStr(1, objName, "Char", vbBinaryCompare) = 0 And InStr(1, Right(objName, 6), "_", vbBinaryCompare) > 0 Then
                    Set fLib = fsLib.Item("Libi8_n_Outputs-DecDatClsX_Tooltip.pdl") 'for act val or ref val objects, with or w/o main attribute
                    Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                Else
                    If (InStr(1, objName, "DatCls2_ActValChar", vbBinaryCompare) Or InStr(1, objName, "DatCls2_RefValChar", vbBinaryCompare)) > 0 Then
                        Set fLib = fsLib.Item("Libi8_n_Outputs-CharDatCls2_Tooltip.pdl") 'for act val or ref val objects, with or w/o main attribute
                        Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                    End If
                End If 'end checking if char exists or not
                If InStr(1, Right(objName, 6), "_", vbBinaryCompare) = 0 Then
                    Set fLib = fsLib.Item("Libi8_n_xTypical-Collection_1_Tooltip+UAL.pdl") 'for act val or ref val objects, with or w/o main attribute
                    Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                End If
                temp = objName
                If InStr(1, Right(temp, 6), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 4), "_", vbBinaryCompare) > 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
            ElseIf (InStr(1, objName, "DatCls3_RefValChar", vbBinaryCompare) Or InStr(1, objName, "DatCls3_ActValChar", vbBinaryCompare)) > 0 And InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then
                Set fLib = fsLib.Item("Libi8_n_Outputs-CharDatCls3_Tooltip.pdl") 'for act val or ref val objects, with or w/o main attribute
                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
            ElseIf InStr(1, objName, "@V3_SMS_SL_", vbBinaryCompare) > 0 Then
                Set fLib = fsLib.Item("Libi8_n_SingleLineObjects_Tooltip.pdl") 'for act val or ref val objects, with or w/o main attribute
                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
            ElseIf InStr(1, objName, "@V3_SMS_Unit", vbBinaryCompare) Then
                Set fLib = fsLib.Item("Libi8_n_Outputs-DecDatClsX_Tooltip.pdl") 'for act val or ref val objects, with or w/o main attribute
                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
                Set objNew = Application.Documents.Item(fLib).HMIObjects(temp & "_1")
                objNew.Selected = True
                If obj.Properties("OutputValue").DynamicStateType = hmiDynamicStateTypeScript Then
                    Set objVBScript = obj.Properties("OutputValue").Dynamic
                    With objVBScript
                        If InStr(1, .sourceCode, "@NOP::#UChg", vbBinaryCompare) > 0 Then
                            tagUChg = Right(.sourceCode, Len(.sourceCode) - InStr(1, .sourceCode, """@NOP::#UChg", vbBinaryCompare))
                            tagUChg = Left(tagUChg, InStr(1, tagUChg, "Unit", vbBinaryCompare) - 1 + Len("Unit"))
                        End If
                    End With
                ElseIf obj.Properties("OutputValue").DynamicStateType = hmiDynamicStateTypeVariableDirect Then
                    Set objVarTrigger = obj.Properties("OutputValue").Dynamic
                    With objVarTrigger
                        tagUChg = .varName
                    End With
                End If
                Set objVBScript = objNew.Properties("ToolTipText").Dynamic
                With objVBScript
                    dynSourceCode = .sourceCode
                    dynSourceCode = Replace(dynSourceCode, "@NOP::#UChg_Default.Unit", tagUChg, , , vbBinaryCompare)
                    trigVarName = .Trigger.VariableTriggers(1).varName
                End With
                Set objVBScriptDest = obj.Properties("ToolTipText").CreateDynamic(hmiDynamicCreationTypeVBScript, "")
                With objVBScriptDest
                    .Trigger.Type = hmiTriggerTypeVariable
                    .Trigger.VariableTriggers.Add "#DskTop.ViewOption_25", hmiVariableCycleTypeOnChange
                    .sourceCode = dynSourceCode
                End With
                Set objVarTrigger = obj.Properties("OutputValue").CreateDynamic(hmiDynamicCreationTypeVariableDirect, tagUChg)
                With objVarTrigger
                    .varName = tagUChg
                    .CycleType = hmiVariableCycleType_500ms
                End With
                tooltipAdded = tooltipAdded + 1
                objNew.Selected = False
                GoTo nextObject
            ElseIf InStr(1, objName, "Unit_Met", vbBinaryCompare) And InStr(1, objName, "Unit_Met_Imp", vbBinaryCompare) = 0 Then
                Set fLib = fsLib.Item("Libi8_n_Devices-Sensors_Tooltip+UAL.pdl") 'for act val or ref val objects, with or w/o main attribute
                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
                Set objNew = Application.Documents.Item(fLib).HMIObjects(temp & "_17")
                objNew.Selected = True
                If obj.Properties("OutputValue").DynamicStateType = hmiDynamicStateTypeScript Then
                    Set objVBScript = obj.Properties("OutputValue").Dynamic
                    With objVBScript
                        If InStr(1, .sourceCode, "@NOP::#UChg", vbBinaryCompare) > 0 Then
                            tagUChg = Right(.sourceCode, Len(.sourceCode) - InStr(1, .sourceCode, """@NOP::#UChg", vbBinaryCompare))
                            tagUChg = Left(tagUChg, InStr(1, tagUChg, "Unit", vbBinaryCompare) - 1 + Len("Unit"))
                        End If
                    End With
                ElseIf obj.Properties("OutputValue").DynamicStateType = hmiDynamicStateTypeVariableDirect Then
                    Set objVarTrigger = obj.Properties("OutputValue").Dynamic
                    With objVarTrigger
                        tagUChg = .varName
                    End With
                End If
                Set objVBScript = objNew.Properties("ToolTipText").Dynamic
                With objVBScript
                    dynSourceCode = .sourceCode
                    dynSourceCode = Replace(dynSourceCode, "@NOP::#UChg_Default.Unit", tagUChg, , , vbBinaryCompare)
                    trigVarName = .Trigger.VariableTriggers(1).varName
                End With
                Set objVBScriptDest = obj.Properties("ToolTipText").CreateDynamic(hmiDynamicCreationTypeVBScript, "")
                With objVBScriptDest
                    .Trigger.Type = hmiTriggerTypeVariable
                    .Trigger.VariableTriggers.Add "#DskTop.ViewOption_25", hmiVariableCycleTypeOnChange
                    .sourceCode = dynSourceCode
                End With
                Set objVarTrigger = obj.Properties("OutputValue").CreateDynamic(hmiDynamicCreationTypeVariableDirect, tagUChg)
                With objVarTrigger
                    .varName = tagUChg
                    .CycleType = hmiVariableCycleType_500ms
                End With
                tooltipAdded = tooltipAdded + 1
                objNew.Selected = False
                GoTo nextObject
            ElseIf InStr(1, objName, "@V3_SMS_Pb", vbBinaryCompare) > 0 Then
'                Set fLib = fsLib.Item("Libi8_n_Buttons-Action_Tooltip+UAL.pdl") 'for act val or ref val objects, with or w/o main attribute
'                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
'                temp = objName
'                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
'                    If IsNumeric(Right(temp, 3)) Then
'                        temp = Left(temp, Len(temp) - 3)
'                    ElseIf IsNumeric(Right(temp, 2)) Then
'                        temp = Left(temp, Len(temp) - 2)
'                    ElseIf IsNumeric(Right(temp, 1)) Then
'                        temp = Left(temp, Len(temp) - 1)
'                    End If
'                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
'                    temp = StrReverse(temp)
'                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
'                    temp = StrReverse(temp)
'                End If
                
                Dim tagCmd As String
                Set objVBScript = obj.Events("OnLButtonUp").Actions(1)
                With objVBScript
                    If InStr(1, .sourceCode, """SMSD", vbBinaryCompare) > 0 Then
                        tagCmd = Right(.sourceCode, Len(.sourceCode) - InStr(1, .sourceCode, """", vbBinaryCompare))
                        tagCmd = Left(tagCmd, InStr(1, tagCmd, """", vbBinaryCompare) - 1 + Len("""") - 1)
                    End If
                End With
                Set objNew = Application.Documents.Item(fLib).HMIObjects(temp)
                objNew.Selected = True
                Set objVBScript = objNew.Properties("ToolTipText").Dynamic
                With objVBScript
                    dynSourceCode = .sourceCode
                    trigVarName = .Trigger.VariableTriggers(1).varName
                End With
                Set objVBScriptDest = obj.Properties("ToolTipText").CreateDynamic(hmiDynamicCreationTypeVBScript, "")
                With objVBScriptDest
                    .Trigger.Type = hmiTriggerTypeVariable
                    .Trigger.VariableTriggers.Add "#DskTop.ViewOption_25", hmiVariableCycleTypeOnChange
                End With
                tooltipAdded = tooltipAdded + 1
                objNew.Selected = False
                GoTo nextObject
            ElseIf (InStr(1, objName, "StdDev_Mot", vbBinaryCompare) Or InStr(1, objName, "NonStd_Mot", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Pmp", vbBinaryCompare) Or InStr(1, objName, "NonStd_Pmp", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Heat", vbBinaryCompare) Or InStr(1, objName, "NonStd_Heat", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Rectifier", vbBinaryCompare) Or InStr(1, objName, "NonStd_Rectifier", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Gen", vbBinaryCompare) Or InStr(1, objName, "NonStd_Gen", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Fan", vbBinaryCompare) Or InStr(1, objName, "NonStd_Fan", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Brake", vbBinaryCompare) Or InStr(1, objName, "NonStd_Brake", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Cooler", vbBinaryCompare) Or InStr(1, objName, "NonStd_Cooler", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_Converter", vbBinaryCompare) Or InStr(1, objName, "NonStd_Converter", vbBinaryCompare) Or _
                InStr(1, objName, "StdDev_VibFeeder", vbBinaryCompare) Or InStr(1, objName, "NonStd_VibFeeder", vbBinaryCompare) Or _
                InStr(1, objName, "@V3_SMS_Pmp", vbBinaryCompare) Or InStr(1, objName, "@V3_SMS_Brake", vbBinaryCompare) Or _
                InStr(1, objName, "@V3_SMS_Cooler", vbBinaryCompare) Or InStr(1, objName, "@V3_SMS_Fan", vbBinaryCompare) Or _
                InStr(1, objName, "@V3_SMS_FrcRls", vbBinaryCompare) Or InStr(1, objName, "@V3_SMS_Gen", vbBinaryCompare) Or _
                InStr(1, objName, "@V3_SMS_Heat", vbBinaryCompare) Or InStr(1, objName, "@V3_SMS_Rectifier", vbBinaryCompare) Or _
                InStr(1, objName, "Movement", vbBinaryCompare) Or InStr(1, objName, "@V3_SMS_Mot_", vbBinaryCompare) Or _
                InStr(1, objName, "Mode", vbBinaryCompare) Or InStr(1, objName, "FrcRls", vbBinaryCompare)) > 0 And _
                InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then
                'check if object is PID developed with SWAT
                Dim tagECUMnAt As String
                tagECUMnAt = ""
                If InStr(1, objName, "@V3_SMS_StdDev_Mot_Fph", vbBinaryCompare) > 0 Then
                    For i = 1 To obj.Properties.Count
                        If obj.Properties(i).DynamicStateType = hmiDynamicStateTypeDynamicDialog Then
                            Set dynamicDialog = obj.Properties.Item(i).Dynamic 'orig object
                            With dynamicDialog
                                trigCount = .Trigger.VariableTriggers.Count 'trigger and source code
                                trigVarName = .Trigger.VariableTriggers(1).varName
                                If InStr(1, trigVarName, "_ECU", vbBinaryCompare) > 0 Then
                                oFile.WriteLine f.Name & ", " & objName & ", because the object is a PID device with ECU tag"
                                    tagECUMnAt = trigVarName
                                    obj.Selected = True
                                    Exit For
                                End If
                            End With
                        End If
                    Next
                End If 'end check if object is PID
                
                If InStr(1, objName, "@V3_SMS_StdDev_Mot_Fph", vbBinaryCompare) > 0 Then
                    If tagECUMnAt <> "" Then
                        Dim objPID As HMICustomizedObject
                        Set objPID = objs(objName)
                        Debug.Print objPID.HMIUdoObjects.Item(1).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(2).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(3).ObjectName & vbCrLf & _
                        objPID.HMIUdoObjects.Item(4).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(5).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(6).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(7).ObjectName & vbCrLf & _
                        objPID.HMIUdoObjects.Item(8).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(9).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(10).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(11).ObjectName & vbCrLf & _
                        objPID.HMIUdoObjects.Item(12).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(13).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(14).ObjectName & vbCrLf & objPID.HMIUdoObjects.Item(15).ObjectName & vbCrLf
                        Rem get MnAt tag for inserting into sourcecode
                        Set objVBScriptDest = obj.Properties("ToolTipText").CreateDynamic(hmiDynamicCreationTypeVBScript, "")
                        With objVBScriptDest
                            .Trigger.Type = hmiTriggerTypeVariable
                            .Trigger.VariableTriggers.Add "#DskTop.ViewOption_25", hmiVariableCycleTypeOnChange
                            .sourceCode = "'Generated using VBA Macro V3.2  Date: 03.03.2018" & vbCrLf _
                            & "' WINCC:TAGNAME_SECTION_START" & vbCrLf _
                            & "Const TagName = " & """" & tagECUMnAt & """" & "" & vbCrLf _
                            & "' WINCC:TAGNAME_SECTION_END" & vbCrLf _
                            & "ToolTip_Trigger = rtUIToolTipObj(TagName) & ""\n"" & rtUIUALMap(Item)"
                        End With
                        tooltipAdded = tooltipAdded + 1
                        GoTo nextObject
                        Else: GoTo nextObject
                    End If
                End If 'end
                
                Set fLib = fsLib.Item("Libi8_n_Devices-Drives_Tooltip+UAL.pdl") 'for act val or ref val objects, with or w/o main attribute
                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
                If InStr(1, objName, "Movement", vbBinaryCompare) > 0 Then
                    temp = temp & "_"
                End If
            ElseIf (InStr(1, objName, "Flap", vbBinaryCompare) Or InStr(1, objName, "Valve", vbBinaryCompare)) > 0 And InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then
                Set fLib = fsLib.Item("Libi8_n_Devices-Valves_Tooltip+UAL.pdl") 'for act val or ref val objects, with or w/o main attribute
                Application.Documents.Open fLib, hmiOpenDocumentTypeVisible
                temp = objName
                If InStr(1, Right(temp, 4), "_", vbBinaryCompare) = 0 Then
                    If IsNumeric(Right(temp, 3)) Then
                        temp = Left(temp, Len(temp) - 3)
                    ElseIf IsNumeric(Right(temp, 2)) Then
                        temp = Left(temp, Len(temp) - 2)
                    ElseIf IsNumeric(Right(temp, 1)) Then
                        temp = Left(temp, Len(temp) - 1)
                    End If
                ElseIf InStr(1, Right(temp, 1), "_", vbBinaryCompare) = 0 Then
                    temp = StrReverse(temp)
                    temp = Right(temp, Len(temp) - InStr(1, temp, "_"))
                    temp = StrReverse(temp)
                End If
            Else
                If InStr(1, objName, "@V3_SMS_", vbBinaryCompare) > 0 Then  'And InStr(1, objName, "_Diag_FPH", vbbinarycompare) = 0 And InStr(1, objName, "_SclVal", vbbinarycompare) = 0
                    oFile.WriteLine f.Name & ", " & objName & ", because the library wasn't explicitly defined for this type of object"
                End If
                GoTo nextObject
            End If
            
performOps:
                
Rem ** END LIBRARY DEFINITIONS ******************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
            
            
Rem *********************************************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
Rem ** START FINDING CORRESPONDING LIBRARY OBJECT ***********************************************************************************************************************************************
            checkBit = 0
            ctrArr = 0
            objLibString = ","
            idxLibArr = 0
            
            Set objLib = Nothing
            
            Do Until checkBit = 1
                
                tagChecker = 0
                tagCount = 0
                objWasChk = 0
                
                If ctrArr > 3 Then
                    MsgBox "Stuck in an infinite loop! Debug or Grafexe will crash! Can't find matching library object."
                    oFile.WriteLine f.Name & ", " & objName & ", stuck in an infinite loop for this object at, " & Now()
                    GoTo nextObject
                End If
                
                libObjScanned = libObjScanned + 1 'that is library objects scanned
                
                If ctrArr > 0 Then
                    objLibString = objLibString & objLib.ObjectName & ","
                End If
                
                Set objsLib = Application.Documents.Item(fLib).HMIObjects
                
                For Each objLib In objsLib
                    If InStr(1, objLib.ObjectName, temp, vbBinaryCompare) Then
                        If InStr(1, objLibString, "," & objLib.ObjectName & ",", vbBinaryCompare) > 0 Then
                            objWasChk = 1
                            Else: Set objLib = objsLib(objLib.ObjectName)
                            Exit For
                        End If
                    End If
                Next
                
                objLib.Selected = True
                
                tagAct = "": tagAdAt = "": tagClAt = "": tagFrcRls = "": tagText = "": tagPnAt = "": tagPpAt = "": tagPV = "": tagLMN = "": tagS1At = "": tagS2At = "": tagS3At = "": tagS4At = "": tagActV = "": tagSetV = "": tagOutV = "": tagMnAt = "": tagMsAt = "": tagOpMask = "": tagActPos = "": tagUnitFactor = "": tagUnitFormat = "": tagUnitOffset = ""
                libTagAdAt = "": libTagActV = "": libTagSetV = "": libTagOutV = "": libTagMnAt = "": libTagMsAt = "": libTagOpMask = "": libTagActPos = "": libTagClAt = "": libTagFrcRls = "": libTagPnAt = "": libTagPpAt = "": libTagPV = "": libTagLMN = "": libTagS1At = "": libTagS2At = "": libTagS3At = "": libTagS4At = "": libTagUnitFactor = "": libTagUnitFormat = "": libTagUnitOffset = "": libTagText = ""
                sTagAdAt = "": sTagClAt = "": sTagFrcRls = "": sTagText = "": sTagPnAt = "": sTagPpAt = "": sTagPV = "": sTagLMN = "": sTagS1At = "": sTagS2At = "": sTagS3At = "": sTagS4At = "": sTagActV = "": sTagSetV = "": sTagOutV = "": sTagMnAt = "": sTagMsAt = "": sTagOpMask = "": sTagActPos = "": sTagUnitFactor = "": sTagUnitFormat = "": sTagUnitOffset = "": tagUnitOffset = "": tagUnitFormat = "": tagUnitFactor = ""
                tagMtrPlate = ""
                trigVarName = ""
                
'                LibTagList = libTagAdAt & vbCrLf & libTagActV & vbCrLf & libTagSetV & vbCrLf & libTagOutV & vbCrLf & libTagMnAt & vbCrLf & libTagMsAt & vbCrLf & libTagOpMask & vbCrLf & libTagActPos & vbCrLf & libTagClAt & vbCrLf & libTagFrcRls & vbCrLf & libTagPnAt & vbCrLf & libTagPpAt & vbCrLf & libTagPV & vbCrLf & libTagLMN & vbCrLf & libTagS1At & vbCrLf & libTagS2At & vbCrLf & libTagS3At & vbCrLf & libTagS4At & vbCrLf & libTagUnitFactor & vbCrLf & libTagUnitFormat & vbCrLf & libTagUnitOffset & vbCrLf & libTagText
'                tagList = tagAdAt & vbCrLf & tagActV & vbCrLf & tagSetV & vbCrLf & tagOutV & vbCrLf & tagMnAt & vbCrLf & tagMsAt & vbCrLf & tagOpMask & vbCrLf & tagActPos & vbCrLf & tagClAt & vbCrLf & tagFrcRls & vbCrLf & tagPnAt & vbCrLf & tagPpAt & vbCrLf & tagPV & vbCrLf & tagLMN & vbCrLf & tagS1At & vbCrLf & tagS2At & vbCrLf & tagS3At & vbCrLf & tagS4At & vbCrLf & tagUnitFactor & vbCrLf & tagUnitFormat & vbCrLf & tagUnitOffset & vbCrLf & tagText & tagMtrPlate
'                tagList = tagAdAt & vbCrLf & tagAct & vbCrLf & tagClAt & vbCrLf & tagFrcRls & vbCrLf & tagText & vbCrLf & tagPnAt & vbCrLf & tagPpAt & vbCrLf & tagPV & vbCrLf & tagLMN & vbCrLf & tagS1At & vbCrLf & tagS2At & vbCrLf & tagS3At & vbCrLf & tagS4At & vbCrLf & tagActV & vbCrLf & tagSetV & vbCrLf & tagOutV & vbCrLf & tagMnAt & vbCrLf & tagMsAt & vbCrLf & tagOpMask & vbCrLf & tagActPos & vbCrLf & tagUnitFactor & vbCrLf & tagUnitFormat & vbCrLf & tagUnitOffset & vbCrLf & tagStatus & vbCrLf & tagCalculateFb & vbCrLf & tagStatusFb & vbCrLf & tagMlcCtrl & vbCrLf & tagMlcLevel & vbCrLf & tagTabNo & vbCrLf & tagMin & vbCrLf & tagMax & vbCrLf & tagRef & vbCrLf & tagHsa & vbCrLf & tagHmo & vbCrLf & tagMlc & vbCrLf & tagRam & vbCrLf & tagFlow & vbCrLf & tagLength & vbCrLf & tagPosition & vbCrLf & tagCounter & vbCrLf & tagMtrPlate
'                MsgBox LibTagList
'                MsgBox tagList

Rem *********** TAG LIST ARRAYS WITH MSGBOX CHECKS ******************************************************

                Dim tagInfoObj
                Dim tagInfoObjLib
                
                Dim key As Variant
                Dim variable As Variant
                Dim member1 As Variant
                Dim member2 As Variant
                Dim member As Variant
                Dim tagObj As New Collection
                Dim tagObjLib As New Collection
                
                Dim objKeys As New Dictionary
                Dim objUniqueTags As Variant
                Dim libKeys As New Dictionary
                Dim libUniqueTags As Variant
                
                Rem new singling out the variables ------
                Dim objServerTags As New Collection
                Dim objUChgTags As New Collection
                Dim objDskTopTags As New Collection
                
                Dim libServerTags As New Collection
                Dim libUChgTags As New Collection
                Dim libDskTopTags As New Collection
                Rem -------------------------------------
                
                Dim servertagslist As String
                Dim libServerTagsList As String
                servertagslist = ""
                libServerTagsList = ""
                
                Dim prefix As Variant
                Dim newPrefix As String
                Dim libDskTop, objDskTop
                Dim libUChg, objUChg
                
                Dim check As Integer
                check = objKeys.Count
                check = objUniqueTags.Count
                check = libKeys.Count
                check = libUniqueTags.Count
                
                Dim lenLib As Integer
                lenLib = 0
                Dim lenObj As Integer
                lenObj = 0
                
                libDskTop = 0
                objDskTop = 0
                libUChg = 0
                objUChg = 0
                
                Set libDskTopTags = Nothing
                Set libServerTags = Nothing
                Set libUChgTags = Nothing
                Set objDskTopTags = Nothing
                Set objServerTags = Nothing
                Set objUChgTags = Nothing
                Set tagObj = Nothing
                Set tagObjLib = Nothing
                Set objKeys = Nothing
                Set objUniqueTags = Nothing
                Set libKeys = Nothing
                Set libUniqueTags = Nothing
                Set tagInfoObj = GetTagInfo(obj)
                Set tagInfoObjLib = GetTagInfo(objLib)
                
                objKeys.RemoveAll
                libKeys.RemoveAll
                
                newPrefix = "SMS_CSPCA" 'new prefix
                For Each variable In tagInfoObj.Items
                  If InStr(1, variable, "::", vbBinaryCompare) And InStr(1, variable, "SMS", vbBinaryCompare) Then
                    prefix = Split(variable, ":", , vbBinaryCompare)
                    prefix = prefix(0)
                  End If
                  tagObj.Add variable
                  objKeys.Add variable, ""
                Next
                
                For Each variable In tagInfoObjLib.Items
                  tagObjLib.Add variable
                  libKeys.Add variable, ""
                Next
                

                objUniqueTags = objKeys.Keys
                libUniqueTags = libKeys.Keys
                lenLib = UBound(libUniqueTags)
                lenObj = UBound(objUniqueTags)
                
                For i = LBound(objUniqueTags) To UBound(objUniqueTags)
'                  If InStr(1, objUniqueTags(i), "UChg", vbbinarycompare) And InStr(1, libUniqueTags(i), "UChg", vbbinarycompare) = 0 Then
'                    MsgBox "not right"
'                  ElseIf InStr(1, objUniqueTags(i), "DskTop", vbbinarycompare) And InStr(1, libUniqueTags(i), "DskTop", vbbinarycompare) = 0 Then
'                    MsgBox "not right"
'                  ElseIf InStr(1, objUniqueTags(i), "NOP", vbbinarycompare) And InStr(1, libUniqueTags(i), "NOP", vbbinarycompare) = 0 Then
'                    MsgBox "not right"
'                  End If
                  If InStr(1, objUniqueTags(i), "DskTop", vbBinaryCompare) Then
                    objDskTop = objDskTop + 1
                    objDskTopTags.Add objUniqueTags(i)
                  End If
                  If InStr(1, objUniqueTags(i), "UChg", vbBinaryCompare) Then
                    objUChg = objUChg + 1
                    objUChgTags.Add objUniqueTags(i)
                  End If
                  If InStr(1, objUniqueTags(i), "DskTop", vbBinaryCompare) = 0 And InStr(1, objUniqueTags(i), "UChg", vbBinaryCompare) = 0 And InStr(1, objUniqueTags(i), "NOP", vbBinaryCompare) = 0 Then
                    objServerTags.Add objUniqueTags(i)
                    servertagslist = servertagslist & objUniqueTags(i) & ", "
                  End If
                Next
                
                For i = LBound(libUniqueTags) To UBound(libUniqueTags)
                  If InStr(1, libUniqueTags(i), "DskTop", vbBinaryCompare) Then
                    libDskTop = libDskTop + 1
                    libDskTopTags.Add libUniqueTags(i)
                  End If
                  If InStr(1, libUniqueTags(i), "UChg", vbBinaryCompare) Then
                    libUChg = libUChg + 1
                    libUChgTags.Add libUniqueTags(i)
                  End If
                  If InStr(1, libUniqueTags(i), "DskTop", vbBinaryCompare) = 0 And InStr(1, libUniqueTags(i), "UChg", vbBinaryCompare) = 0 And InStr(1, libUniqueTags(i), "NOP", vbBinaryCompare) = 0 Then
                    libServerTags.Add libUniqueTags(i)
                    libServerTagsList = libServerTagsList & libUniqueTags(i) & ", "
                  End If
                Next
                
'                If lenLib <> lenObj Then
'                  If libDskTop <> objDskTop Then
'                    If (Abs(lenLib - lenObj) - Abs(libDskTop - objDskTop) <> 0) Then
'                      'MsgBox "one has more tags than the other, check"
'                      oFile.WriteLine f.Name & ", " & objName & ", " & objLib.ObjectName & ", 0 s" & ", " & libServerTagsList & ", " & servertagslist & ", was not exchanged because of different server tag count check"
'
'                      Set libDskTopTags = Nothing
'                      Set libServerTags = Nothing
'                      Set libUChgTags = Nothing
'                      Set objDskTopTags = Nothing
'                      Set objServerTags = Nothing
'                      Set objUChgTags = Nothing
'                      Set tagObj = Nothing
'                      Set tagObjLib = Nothing
'                      Set objKeys = Nothing
'                      Set objUniqueTags = Nothing
'                      Set libKeys = Nothing
'                      Set libUniqueTags = Nothing
'
'                      'GoTo nextObject
'                    End If
'                  ElseIf libUChg <> objUChg Then
'                    If (Abs(lenLib - lenObj) - Abs(libUChg - objUChg) <> 0) Then
'                      MsgBox "one has more tags than the other, check"
'                    End If
'                  End If
'                End If
                
                If objServerTags.Count <> libServerTags.Count Then
                  'MsgBox "different server tag count, check"
                  oFile.WriteLine f.Name & ", " & objName & ", " & objLib.ObjectName & ", 0 s" & ", " & libServerTagsList & ", " & servertagslist & ", was not exchanged because of different server tag count check"
                  
                  Set libDskTopTags = Nothing
                  Set libServerTags = Nothing
                  Set libUChgTags = Nothing
                  Set objDskTopTags = Nothing
                  Set objServerTags = Nothing
                  Set objUChgTags = Nothing
                  Set tagObj = Nothing
                  Set tagObjLib = Nothing
                  Set objKeys = Nothing
                  Set objUniqueTags = Nothing
                  Set libKeys = Nothing
                  Set libUniqueTags = Nothing
                  
                  'GoTo nextObject
                Else 'lib object is ok for what we need
                checkBit = 1
                End If
                
Rem *********** TAG LIST ARRAYS *******************************************************************
nextLibObj:
                    'checkBit = 1 'uncomment to override library object choice checker, select first found
            Loop
            
Rem *********************************************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
Rem ** END OF FINDING CORRESPONDING LIBRARY OBJECT **********************************************************************************************************************************************
            
            
Rem ** TAKING CHOSEN LIBRARY OBJECT, DELETING OLD OBJECTS AND BRINGING NEW OBJECTS IN ***********************************************************************************************************
Application.Documents(f).selection.DeselectAll
Application.Documents(fLib).selection.DeselectAll
If objLib.ObjectName = "" Then
    oFile.WriteLine f.Name & ", " & objName & ", because a corresponding library object was not found"
    GoTo nextObject
Else
    objLib.Selected = True
End If
Application.Documents(fLib).selection.CopySelection
Application.Documents.Open f, hmiOpenDocumentTypeVisible
Application.Documents(f).PasteClipboard
Set objNew = Application.Documents(f).selection(1) ' settle object
Set objNew = Application.Documents(f).HMIObjects.Item(objNew.ObjectName)
objNew.Top = obj.Top
objNew.Left = obj.Left
objNew.Layer = obj.Layer
If InStr(1, objNew.ObjectName, "_ClAt", vbBinaryCompare) > 0 Then
    objNew.Height = obj.Height
    objNew.Width = obj.Width
End If

obj.Delete 'delete old object
Rem ** END TAKING CHOSEN LIBRARY OBJECT, DELETING OLD OBJECTS AND BRINGING NEW OBJECTS IN ******************************************************************************************************

            
Rem *********************************************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
Rem ** REPLACING TAGS IN NEWLY BROUGHT OBJECT, ALL TYPES OF DYNAMICS FROM OBTAINED TAG LIST *****************************************************************************************************

Dim oldSrc As String
Dim newSrc As String
Dim propName As String
Dim tag As Variant
Dim idx As Integer
Dim prop As HMIProperty
Dim props As HMIProperties
Dim trig As HMIVariableTrigger
Dim trigs As HMIVariableTriggers


Set props = objNew.Properties

For Each prop In props

  propName = prop.DisplayName & " " & prop.Name
  j = 0
  
    If prop.DynamicStateType = hmiDynamicStateTypeDynamicDialog Then
        Set dynamicDialog = prop.Dynamic
        With dynamicDialog
        
            oldSrc = .sourceCode
            
            idx = 1
            For Each tag In libServerTags
                  objServerTags(idx) = Replace(objServerTags(idx), prefix, newPrefix, , , vbBinaryCompare)
                  .sourceCode = Replace(.sourceCode, tag, objServerTags(idx), , , vbBinaryCompare):
                idx = idx + 1
            Next

            idx = 1
            For Each tag In libUChgTags
                  .sourceCode = Replace(.sourceCode, tag, objUChgTags(idx), , , vbBinaryCompare)
                idx = idx + 1
            Next
            
            newSrc = .sourceCode
            
            If .sourceCode = oldSrc And oldSrc <> newSrc Then
              MsgBox "hasn't changed source code here"
            End If
            
            Set trigs = .Trigger.VariableTriggers
            For Each trig In trigs
              If trig.CycleType = hmiVariableCycleType_2s Then
                trig.CycleType = hmiVariableCycleType_500ms
              End If
              If trig.varName = "" Or Len(trig.varName) < 2 Then
                  MsgBox "trigger was emptied here"
              End If
            Next
            
        End With
    ElseIf prop.DynamicStateType = hmiDynamicStateTypeScript Then
        Set objVBScript = prop.Dynamic
        With objVBScript
        
            oldSrc = .sourceCode
            
            .sourceCode = Replace(.sourceCode, prefix, newPrefix, , , vbBinaryCompare)
            
            idx = 1
            For Each tag In libServerTags
              If InStr(1, .sourceCode, tag, vbBinaryCompare) Then
                  objServerTags(idx) = Replace(objServerTags(idx), prefix, newPrefix, , , vbBinaryCompare)
                  .sourceCode = Replace(.sourceCode, tag, objServerTags(idx), , , vbBinaryCompare)
                idx = idx + 1
              End If
            Next
            
            idx = 1
            For Each tag In libUChgTags
              If InStr(1, .sourceCode, tag, vbBinaryCompare) Then
                  .sourceCode = Replace(.sourceCode, tag, objUChgTags(idx), , , vbBinaryCompare)
                idx = idx + 1
              End If
            Next
            
            newSrc = .sourceCode
            
            If .sourceCode = oldSrc And oldSrc <> newSrc Then
              MsgBox "hasn't changed source code here"
            End If
            
            Set trigs = .Trigger.VariableTriggers
            For Each trig In trigs
              If trig.CycleType = hmiVariableCycleType_2s Then
                trig.CycleType = hmiVariableCycleType_500ms
              End If
              If trig.varName = "" Or Len(trig.varName) < 2 Then
                  MsgBox "trigger was emptied here"
              End If
            Next
            
        End With
    ElseIf prop.DynamicStateType = hmiDynamicStateTypeVariableDirect Or prop.DynamicStateType = hmiDynamicStateTypeVariableIndirect Then
        Set objVarTrigger = prop.Dynamic
        With objVarTrigger
        
            idx = 1
            For Each tag In libServerTags
              If InStr(1, .varName, tag, vbBinaryCompare) Then
                objServerTags(idx) = Replace(objServerTags(idx), prefix, newPrefix, , , vbBinaryCompare)
                .varName = objServerTags(idx)
                idx = idx + 1
              End If
            Next
            
            idx = 1
            For Each tag In libUChgTags
              If InStr(1, .varName, tag, vbBinaryCompare) Then
                .varName = objUChgTags(idx)
                idx = idx + 1
              End If
            Next
            
              If trig.CycleType = hmiVariableCycleType_2s Then
                trig.CycleType = hmiVariableCycleType_500ms
              End If
              
            If Len(.varName) < 2 Then
              MsgBox "trigger was emptied here"
            End If
            
        End With
    End If
Next



objEndTimer = Timer()
objTime = FormatNumber(objEndTimer - objTimer, 2)
oFile.WriteLine f.Name & ", " & objName & ", " & objNew.ObjectName & ", " & objTime & " s" & ", " & libServerTagsList & ", " & servertagslist

Set libDskTopTags = Nothing
Set libServerTags = Nothing
Set libUChgTags = Nothing
Set objDskTopTags = Nothing
Set objServerTags = Nothing
Set objUChgTags = Nothing
Set tagObj = Nothing
Set tagObjLib = Nothing
Set objKeys = Nothing
Set objUniqueTags = Nothing
Set libKeys = Nothing
Set libUniqueTags = Nothing

objsChanged = objsChanged + 1
Application.Documents(fLib).selection.DeselectAll
Application.Documents(f).selection.DeselectAll
            
Rem ** END REPLACING TAGS IN NEWLY BROUGHT OBJECT, ALL TYPES OF DYNAMICS FROM OBTAINED TAG LIST *************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
Rem *********************************************************************************************************************************************************************************************
nextObject:
            
        Next 'next obj in objs in processed page (not lib)
        
        Rem end of script to be performed
        
        Rem check if the file needs to be saved or not, else it will just close it
        If objsChanged > 0 Or tooltipAdded > 0 Then
            
            Application.Documents(f).Save 'As ("C:\Project\lib\" & f.Name & "_proc.pdl")
            oFile.WriteLine f.Name & ", " & objsChanged & ", index of objects changed at " & Now() & "," & f.Name & ", " & tooltipAdded & ", index of tooltips added at " & Now()  '& FormatNumber(Timer() - StartTime, 2) & " seconds, saved file"
        End If
        Set objs = Nothing
        
    End If
    Rem end multiple page filter *************************************************************
    'End If 'end multiple file filter
    'Next   'next row in array of files
    Rem ***************************************************************************************
    
nextFile:
    
    oFile.WriteLine f.Name & ", " & objsChanged & ", index of objects changed at " & Now() & "," & f.Name & ", " & tooltipAdded & ", index of tooltips added at " & Now()  '& FormatNumber(Timer() - StartTime, 2) & " seconds"
    
Next 'next file in graCS

Set fso = Nothing
Set Files = Nothing

endTime = Timer()

oFile.Close

MsgBox objsScanned & " objects have been scanned" _
& vbCrLf & libObjScanned & " library objects have been scanned." _
& vbCrLf & tooltipAdded & " tooltips have been added;" _
& vbCrLf & objsChanged & " objects have been taken from the library in " & FormatNumber(endTime - startTime, 2) & " seconds", vbOKOnly

End Sub
