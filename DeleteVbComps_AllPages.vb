Sub CallMain()

Call RunThroughAllPages

End Sub

Sub DeleteVBComponent()

Dim activeIde As Object
'Ignore errors
On Error Resume Next

'Delete the components
i = 3 '3 is the index of the vbaproject of the pdl

For j = 0 To Application.VBE.VBProjects(i).VBComponents.Count
    Debug.Print Application.VBE.VBProjects(i).VBComponents(j).Type & " " & Application.VBE.VBProjects(i).VBComponents(j).Name
    If Left(Application.VBE.VBProjects(i).VBComponents(j).Name, Len("ThisDocument")) = "ThisDocument" Then
        LineCount = Application.VBE.VBProjects(i).VBComponents(j).CodeModule.CountOfLines
        Debug.Print LineCount
        Application.VBE.VBProjects(i).VBComponents(j).CodeModule.DeleteLines 1, LineCount
    End If
    Application.VBE.VBProjects(i).VBComponents.Remove Application.VBE.VBProjects(i).VBComponents(j)
Next

End Sub

Sub calling_procedure()

Call DeleteVBComponent
Call DeleteVBComponent
Call DeleteVBComponent

End Sub

Sub RunThroughAllPages()

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
Dim pagesChanged As Integer

StartTime = Timer()

Set fso = CreateObject("Scripting.FilesystemObject")

FolderPath = "C:\Processed"

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
Rem rem multiple page filter *********************************************************
'z = 0
'For z = 0 To (UBound(filterPages) - 1)
'If filterPages(z) <> "" And InStr(1, f, filterPages(z), vbTextCompare) > 0 Then 'filter to single out multiple pages, comment if not needed

Rem **********************************************************************************
If InStr(1, f, filterpdl, vbTextCompare) > 0 Then  'filter through .pdl files only
Rem end of file filter

    Application.Documents.Open f, hmiOpenDocumentTypeVisible
    yyyy = f.Name
    
    Dim obj As HMIObject
    Dim objs As HMIObjects
    
    Set objs = ActiveDocument.HMIObjects
        For Each obj In objs 'cycle through objects to check and perform changes
        
        Rem script to be performed on the filtered pages
        
        Call calling_procedure
        
        k = k + 1
        pagesChanged = pagesChanged + 1
        
        Rem end of script to be performed
        
        Next obj
    
    Rem check if the file needs to be saved or not, else it will just close it
    If k > 0 Then
        ActiveDocument.Save
    End If
    Set objs = Nothing
    ActiveDocument.Close
    
End If

Rem end multiple page filter *************************************************************
'End If 'end multiple file filter
'Next   'next row in array of files
Rem ***************************************************************************************


Next 'next file in graCS

Set fso = Nothing
Set Files = Nothing

EndTime = Timer()
MsgBox pagesChanged & " pages have had vb components deleted/cleared in " & FormatNumber(EndTime - StartTime, 2) & " seconds"

End Sub
