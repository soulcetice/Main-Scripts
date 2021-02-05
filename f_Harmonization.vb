Option Explicit

Public issueCount As Integer

Sub buildUserForm()

Dim uf As UserForm
Dim f As Frame
Dim t As Object
Set uf = HarmonizationHelper

Set f = uf.Frame1
Dim c As CheckBox
Set c = uf.CheckBox1

End Sub

Sub showUserForm()
  HarmonizationHelper.Show
End Sub
'''
''' Returns the count of the specified character in the specified string.
'''
Public Function CountChrInString(Expression As String, Character As String) As Long
'
' ? CountChrInString("a/b/c", "/")
'  2
' ? CountChrInString("a/b/c", "\")
'  0
' ? CountChrInString("//////", "/")
'  6
' ? CountChrInString(" a / b / c ", "/")
'  2
' ? CountChrInString("a/b/c", " / ")
'  0
'
    Dim iResult As Long
    Dim sParts() As String

    sParts = Split(Expression, Character)

    iResult = UBound(sParts, 1)

    If (iResult = -1) Then
    iResult = 0
    End If

    CountChrInString = iResult

End Function

Public Function HarmonizeContainers(repair As Boolean) As Integer
  Dim prop As HMIProperty
  Dim obj As HMIObject
  Dim sobj As HMIObject
  Dim ocont As HMIRectangle
  Dim oconts As Collection
  Dim ounit As HMIObject
  Dim ounitText As HMIStaticText
  Dim ounits As New Collection
  Dim otitle As HMIStaticText
  Dim otitles As New Collection
  Dim include, notinclude As String
  Dim foundContainers As New Collection
  Dim foundTitles As New Collection
  Dim c As Variant
  Dim t As New Collection
  Dim ti As HMIRectangle
  Dim i As Integer
  Dim col As HMIObjects
  Dim startTime, endTime
  Dim searchedAreas As New Collection
  Dim area As Variant
  Dim areaLeft, areaTop, areaWidth, areaHeight
  Dim uf As UserForm
  Dim secocont As HMIRectangle
  Dim o As HMIRectangle
  Dim contHighlights As New Collection
  Dim titleHighlights As New Collection
  Dim unitHighlights As New Collection
  Dim offset As Integer
  
  issueCount = 0
  
  startTime = Timer()
  
  Rem restrict searches to relevant objects
  Dim objs As HMIObjects
  Set objs = ActiveDocument.HMIObjects
  Dim otitleSet As New Collection
  Set otitleSet = GetHMIObjectsByType("HMIStaticText", objs)
  Dim ounitSet1 As New Collection
  Dim ounitSet2 As New Collection
  Dim ounitSet3 As New Collection
  Dim ounitSet As New Collection
  Set ounitSet1 = GetHMIObjectsByType("HMIIOField", objs, "@V3_SMS_Unit")
  Set ounitSet2 = GetHMIObjectsByType("HMIIOField", objs, "Unit_Met_Imp")
  Set ounitSet3 = GetHMIObjectsByType("HMIStaticText", objs, , , , , "[", "]")
  Set ounitSet = joinCollections(ounitSet1, ounitSet2)
  Set ounitSet = joinCollections(ounitSet, ounitSet3)
  Dim potentialContainers As New Collection
    
    'If Expression <= 100 Then
    '  BorderColor_Trigger = -2147483527 '121 'class3
    'ElseIf Expression <= 200 Then
    '  BorderColor_Trigger = -2147483561 '87 'class2
        
  Dim oSMS As New Collection 'customized objects
  Set oSMS = GetHMIObjectsByType("HMICustomizedObject", objs, "@V3")
  
  For Each obj In oSMS 'looks to standard object, then at possibly existing container from potential set
    ActiveDocument.Selection.DeselectAll 'clear selection for next steps
    'obj.Selected = True
    
'    Dim oconts As Collection
    Set oconts = getContainers(obj, objs, 0)
    For Each o In oconts
      'o.Selected = True
      On Error Resume Next
      foundContainers.Add o, o.ObjectName
      searchedAreas.Add o
    Next
  Next

  For Each o In foundContainers
    'tryotitleAgain:
      For Each sobj In otitleSet
        If sobj.left < (o.left + o.width) And sobj.left > (o.left - 5) And sobj.top < (o.top + o.height) And sobj.top < (o.top) Then
          If (o.top - sobj.top < (10 + 5)) _
          And (Abs(o.left - sobj.left) < 30) Then
            offset = 0
            If sobj.FONTBOLD = True And sobj.FONTSIZE = 14 Then
              offset = 5
            End If
              Set otitle = sobj
              'otitle.Selected = True
              Rem check if not ok to add
            Dim count As Integer
            count = (CountChrInString(otitle.text, vbCrLf) + 1) * 2
              If (otitle.left <> o.left + 10 - offset) Or left(otitle.text, 1) <> " " Or right(otitle.text, 1) <> " " Or CInt(otitle.top) <> CInt(o.top - otitle.height / count) Then
                otitles.Add otitle
              End If
              Exit For
          End If
        End If
      Next sobj
      If otitle Is Nothing Then
        'MsgBox "trying search for title object again/there's no title for this container"
        'GoTo noTitle
        'GoTo tryotitleAgain
      Else
        If repair = True Then Call ArrangeContainerTitle(otitle, o, offset)
        For Each sobj In ounitSet
          Set ounit = Nothing
          If (Abs(otitle.top - sobj.top) < 5) _
          And (Abs(otitle.left - sobj.left) < o.width - sobj.width) And sobj.left > otitle.left Then
            Set ounit = sobj
            'ounit.Selected = True
            Rem check if not ok to add
            If ounit.Type = "HMIIOField" Then
              If CInt(ounit.left) <> CInt(otitle.left + otitle.width - 2) Or right(ounit.InputValue, 1) <> " " Or left(ounit.InputValue, 1) = " " Or ounit.top <> otitle.top Or ounit.AlignmentLeft <> 0 Or ounit.AdaptBorder <> True Then
                ounits.Add ounit
              Else
                Set ounit = Nothing
              End If
            ElseIf ounit.Type = "HMIStaticText" Then
              If ounit.FONTBOLD <> True Or CInt(ounit.left) <> CInt(otitle.left + otitle.width - 2) Or right(ounit.text, 1) <> " " Or left(ounit.text, 1) = " " Or ounit.top <> otitle.top Or ounit.AlignmentLeft <> 0 Or ounit.AdaptBorder <> True Then
                ounits.Add ounit
              Else
                Set ounit = Nothing
              End If
            End If
          End If
          If Not ounit Is Nothing And repair = True Then
            ActiveDocument.Selection.DeselectAll
            'ounit.Selected = True
            'otitle.Selected = True
            'o.Selected = True
            Call ArrangeContainerTitleUnit(ounit, otitle, o)
          ElseIf ounit Is Nothing Then
            'MsgBox "trying search for ounit object again"
          End If
        Next sobj
      End If
noTitle:
  Next o
  endTime = Timer()
  
  Rem Set contHighlights = CreateHighlights(searchedAreas)
  Set titleHighlights = CreateHighlights(otitles)
  Set unitHighlights = CreateHighlights(ounits)
  
  issueCount = unitHighlights.count + titleHighlights.count
  
  If repair = True Then
    Call DeleteObjectsFoundInCol(contHighlights)
    Call DeleteObjectsFoundInCol(titleHighlights)
    Call DeleteObjectsFoundInCol(unitHighlights)
  End If
  
  Set uf = HarmonizationHelper
  ActiveDocument.Selection.DeselectAll
  On Error Resume Next
  uf.Hide
  
  Dim what As String
  If repair = True Then
    what = "repair of container objects"
  Else
    what = "check of container objects"
  End If
  
  HarmonizeContainers = issueCount
  
  uf.ListBox1.AddItem ("Ran " & what & ", " & issueCount & " issues in " & FormatNumber(endTime - startTime, 3) & " seconds")
End Function

Public Function HarmonizeButtons(repair As Boolean) As Integer

  Dim buttons As New Collection
  Dim versionInfos As New Collection
  Dim modes As New Collection
  Dim leds As New Collection
  Dim sorted As New Collection
  Dim grp As New Collection
  Dim elem As Variant
  Dim elemSec As Variant
  Dim bt As HMIButton
  Dim startTime, endTime
  Dim tops As New Collection
  Dim col As New Collection
  Dim indexLeft
  Dim colMain As New Collection
  Dim colSpaces As New Collection
  Dim avgSpce
  Dim spce
  Dim curVersionInfo As HMIStaticText
  Dim curMode As HMICustomizedObject
  Dim curLed As HMICustomizedObject
  
  Dim i As Integer
  Dim j As Long
  Dim l As Integer
  
  Dim highlightMain As New Collection
  Dim highlightSec As New Collection
  Dim difLeft As New Collection
  Dim minLeft As Integer
  
  Dim lessThanTen As New Collection
  Dim difThanAvg As New Collection
  Dim intAvg
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  
  Dim minimumSpacing: Dim averageSpacing
  minimumSpacing = CInt(uf.TextBox1.value)
  averageSpacing = CInt(uf.TextBox2.value)
  
  issueCount = 0
  
  startTime = Timer()
  
  Set buttons = GetHMIObjectsByType("HMIButton", ActiveDocument.HMIObjects)
  Set versionInfos = GetHMIObjectsByType("HMIStaticText", ActiveDocument.HMIObjects, "VersionInfo")
  Set modes = GetHMIObjectsByType("HMICustomizedObject", ActiveDocument.HMIObjects, "@V3_SMS_Mode_")
  Set leds = GetHMIObjectsByType("HMICustomizedObject", ActiveDocument.HMIObjects, "@V3_SMS_SignLmp_")
    
redoSpacing:
  Set col = getButtonGroups(buttons)
  Set colMain = col(1)
  Set colSpaces = col(2)
  Set col = Nothing
    
  Rem spacing issues find and repair
  For i = 1 To colSpaces.count
    For j = 1 To colSpaces(i).count
    If colSpaces(i)(j).count > 0 Then
      avgSpce = 0
      For Each elem In colSpaces(i)(j)
        avgSpce = avgSpce + elem
      Next
      avgSpce = avgSpce / colSpaces(i)(j).count
      If avgSpce = Int(avgSpce) Then intAvg = avgSpce
      If intAvg = 0 Then intAvg = 14
    End If
      l = 1
      For Each elem In colSpaces(i)(j)
        Rem avg spacing issues
        Dim fixed As Boolean
        If elem <> averageSpacing Then
          On Error Resume Next
          difThanAvg.Add colMain(i)(j)(l + 1), colMain(i)(j)(l + 1).ObjectName 'above this, dif than avg spacing
          colMain(i)(j)(l + 1).Selected = True
          If repair = True Then 'repair top here
          Dim term
          term = 1
            If l = 1 Then
              l = 0: term = -1
            End If
            Set curVersionInfo = GetObjectFromCollectionByCoordinates(versionInfos, colMain(i)(j)(l + 1).left, colMain(i)(j)(l + 1).top, 2, 2)
            Set curMode = GetObjectFromCollectionByCoordinates(modes, colMain(i)(j)(l + 1).left + 5, colMain(i)(j)(l + 1).top + 5, 5, 5)
            Set curLed = GetObjectFromCollectionByCoordinates(leds, colMain(i)(j)(l + 1).left + 26, colMain(i)(j)(l + 1).top + 5, 5, 5)
            colMain(i)(j)(l + 1).top = colMain(i)(j)(l + 1).top - term * (elem - averageSpacing)
            curVersionInfo.top = colMain(i)(j)(l + 1).top
            curMode.top = colMain(i)(j)(l + 1).top + 3
            curLed.top = colMain(i)(j)(l + 1).top + 2
            GoTo redoSpacing
            fixed = True
          End If
        End If
        Rem less than 10 px spacing issues
        If elem < minimumSpacing Then
          On Error Resume Next
          lessThanTen.Add colMain(i)(j)(l + 1), colMain(i)(j)(l + 1).ObjectName 'above this, less than 10
          colMain(i)(j)(l + 1).Selected = True
          If repair = True And fixed = False Then 'repair top here
            Set curVersionInfo = GetObjectFromCollectionByCoordinates(versionInfos, colMain(i)(j)(l + 1).left, colMain(i)(j)(l + 1).top, 2, 2)
            Set curMode = GetObjectFromCollectionByCoordinates(modes, colMain(i)(j)(l + 1).left + 5, colMain(i)(j)(l + 1).top + 5, 5, 5)
            Set curLed = GetObjectFromCollectionByCoordinates(leds, colMain(i)(j)(l + 1).left + 26, colMain(i)(j)(l + 1).top + 5, 5, 5)
            colMain(i)(j)(l + 1).top = colMain(i)(j)(l + 1).top - (elem - minimumSpacing)
            curVersionInfo.top = colMain(i)(j)(l + 1).top
            curMode.top = colMain(i)(j)(l + 1).top + 3
            curLed.top = colMain(i)(j)(l + 1).top + 2
            If l < colSpaces(i)(j).count Then colSpaces(i)(j)(l + 1) = colSpaces(i)(j)(l + 1) - (elem - minimumSpacing)
            elem = minimumSpacing
          End If
        End If
        l = l + 1
      Next
    Next
  Next
    
  issueCount = issueCount + ActiveDocument.Selection.count
  ActiveDocument.Selection.DeselectAll
  
  Rem left alignment issues (wrt leftmost coordinate)
  For i = 1 To colMain.count
    For l = 1 To colMain(i).count
      minLeft = 3000
      Set difLeft = Nothing
      For Each elem In colMain(i)(l)
        If elem.left < minLeft Then minLeft = elem.left
      Next
      j = 1
      For Each elem In colMain(i)(l)
        If elem.left <> minLeft Then
          ActiveDocument.Selection.DeselectAll
          colMain(i)(l)(j).Selected = True
          On Error Resume Next
          difLeft.Add colMain(i)(l)(j), colMain(i)(l)(j).ObjectName.value
          If repair = True Then 'repair left here
            Set curVersionInfo = GetObjectFromCollectionByCoordinates(versionInfos, colMain(i)(j)(l + 1).left, colMain(i)(j)(l + 1).top, 2, 2)
            Set curMode = GetObjectFromCollectionByCoordinates(modes, colMain(i)(j)(l + 1).left + 5, colMain(i)(j)(l + 1).top + 5, 5, 5)
            Set curLed = GetObjectFromCollectionByCoordinates(leds, colMain(i)(j)(l + 1).left + 26, colMain(i)(j)(l + 1).top + 5, 5, 5)
            colMain(i)(l)(j).left = minLeft
            curVersionInfo.left = colMain(i)(l)(j).left
            curMode.left = colMain(i)(l)(j).left + 4
            curLed.left = colMain(i)(l)(j).left + 27
          End If
        End If
        j = j + 1
      Next
      Set highlightSec = CreateLeftHighlights(difLeft, minLeft) 'highlighting left alignment issues compared to min left
      issueCount = issueCount + highlightSec.count
    Next
  Next
  
  ActiveDocument.Selection.DeselectAll
  
  Set highlightMain = CreateHighlights(difThanAvg) 'will only highlight what was moved already
  Set highlightMain = CreateHighlights(lessThanTen) 'will only highlight what was moved already
  
  ActiveDocument.Selection.DeselectAll
  
  endTime = Timer()
  
  Dim what As String
  If repair = True Then
    what = "repair of button groups"
  Else
    what = "check of button groups"
  End If
  
  HarmonizeButtons = issueCount
  
  uf.ListBox1.AddItem ("Ran " & what & ", " & issueCount & " issues in " & FormatNumber(endTime - startTime, 3) & " seconds")
End Function

Public Function HarmonizeDC3Spacing(repair As Boolean) As Integer
'Vertical DC 3
'? distance for objects to left side = 20 pixel on left side
'? distance to top and bottom = 10 pixel

'Horizontal DC 3
'? distance for objects to left side = 10 pixel
'? distance to top and bottom = 10 pixel
'? distance between bar graph containers = 10 pixel

Dim vertical As Boolean

  If ActiveDocument.width > ActiveDocument.height Then
    vertical = False
    Else
    vertical = True
    End If
    
    
Dim colSearchResults As HMICollection
Dim barGraphContainers As HMICollection
Dim issuesObjects As New Collection
Dim LeftIssuesObjects As New Collection
Dim TopIssuesObjects As New Collection
Dim BottomIssuesObjects As New Collection
Dim objMember As HMIObject
Dim furtherFiltering As New Collection
'

Set colSearchResults = ActiveDocument.HMIObjects.Find(PropertyName:="Left")
For Each objMember In colSearchResults
  If objMember.Layer < 29 Then 'And objMember.Type <> "HMIRectangle" Then
    furtherFiltering.Add objMember, objMember.ObjectName
  End If
Next

Set barGraphContainers = ActiveDocument.HMIObjects.Find(ObjectType:="HMIRectangle")
    For Each objMember In barGraphContainers
      
    Next
    
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  
  If vertical = True Then
    For Each objMember In furtherFiltering
      If objMember.left < uf.dc3_left.text Then
        LeftIssuesObjects.Add objMember, objMember.ObjectName
        End If
      If objMember.top < uf.dc3_topbottom.text Then
        TopIssuesObjects.Add objMember, objMember.ObjectName
        End If
      If objMember.top + objMember.height > ActiveDocument.height - uf.dc3_topbottom.text Then
        BottomIssuesObjects.Add objMember, objMember.ObjectName
        End If
    Next
  ElseIf vertical = False Then
    For Each objMember In furtherFiltering
      If objMember.left < uf.dc3_horiz_left.text Then
        LeftIssuesObjects.Add objMember, objMember.ObjectName
        End If
      If objMember.top < uf.dc3_horiz_topbottom.text Then
        TopIssuesObjects.Add objMember, objMember.ObjectName
        End If
      If objMember.top + objMember.height > ActiveDocument.height - 10 Then
        BottomIssuesObjects.Add objMember, objMember.ObjectName
        End If
    Next
    'to do spacing between bargraph containers
  End If
  
  If repair = False Then
    Dim highlightLeft
    Dim highlightTop
    Dim highlightBottom
    Dim highlightMain
    Set issuesObjects = joinCollections(LeftIssuesObjects, TopIssuesObjects)
    Set issuesObjects = joinCollections(issuesObjects, BottomIssuesObjects)
    Set highlightMain = CreateHighlights(issuesObjects, "DC3 Spacing Issue") 'will only highlight what was moved already
  End If
  
  Debug.Print issuesObjects.count

End Function


Public Function HarmonizePopups(repair As Boolean) As Integer

Dim bottomLeft As HMIObject
Dim bottomRight As HMIObject
Dim shoulderLeft As HMIObject
Dim shoulderRight As HMIObject
Dim titleText As HMIObject
Dim headerRight As HMIObject
Dim headerLeft As HMIObject
Dim objs As HMIObjects
Dim count As Integer
Dim issuesObjects As New Collection

  Set objs = ActiveDocument.HMIObjects
  
  If InStr(1, ActiveDocument.name, "_p_", vbBinaryCompare) = 0 Then
    Debug.Print "This was not a popup, got out of execution"
    Exit Function
  End If
  
  'change to use find function in objs....
  On Error Resume Next
  Set bottomLeft = objs("@BottomLeft")
  On Error Resume Next
  Set bottomRight = objs("@BottomRight")
  On Error Resume Next
  Set shoulderLeft = objs("@ShoulderLeft")
  On Error Resume Next
  Set shoulderRight = objs("@ShoulderRight")
  On Error Resume Next
  Set headerLeft = objs("@HeaderLeft")
  On Error Resume Next
  Set headerRight = objs("@HeaderRight")
  On Error Resume Next
  Set titleText = objs("@TitleText")
  
  Dim width
  Dim height
  width = ActiveDocument.width
  height = ActiveDocument.height
  
  If Not bottomLeft Is Nothing Then
    If bottomLeft.left <> 0 Or bottomLeft.top <> height - bottomLeft.height - 1 Then
      If repair = True Then
        bottomLeft.left = 0
        bottomLeft.top = height - bottomLeft.height - 1
      Else
        issuesObjects.Add bottomLeft, bottomLeft.ObjectName
      End If
      count = count + 1
      End If
    
  End If
  
  If Not bottomRight Is Nothing Then
    If bottomRight.left <> width - bottomRight.width - 1 Or bottomRight.top <> height - bottomRight.height - 1 Then
      If repair = True Then
        bottomRight.left = width - bottomRight.width - 1
        bottomRight.top = height - bottomRight.height - 1
      Else
        issuesObjects.Add bottomRight, bottomRight.ObjectName
      End If
      count = count + 1
      End If
    
  End If
  
  If Not shoulderRight Is Nothing Then
    If shoulderRight.left <> width - shoulderRight.width - 1 Or shoulderRight.top <> 0 Then
      If repair = True Then
        shoulderRight.top = 0
        shoulderRight.left = width - shoulderRight.width - 1
      Else
        issuesObjects.Add shoulderRight, shoulderRight.ObjectName
      End If
      count = count + 1
      End If
    
  End If
  
  If Not shoulderLeft Is Nothing Then
    If shoulderLeft.left <> 0 Or shoulderLeft.top <> 0 Then
      If repair = True Then
        shoulderLeft.top = 0
        shoulderLeft.left = 0
      Else
        issuesObjects.Add shoulderLeft, shoulderLeft.ObjectName
      End If
      count = count + 1
      End If
    
  End If
  
  If Not titleText Is Nothing Then
    If titleText.left <> 6 Or titleText.top <> 1 Then
      If repair = True Then
        titleText.top = 1
        titleText.left = 6
      Else
        issuesObjects.Add titleText, titleText.ObjectName
      End If
      count = count + 1
      End If
    
  End If
  
'  headerLeft.Top = 0
'  headerLeft.Left = 0
  
  If Not headerRight Is Nothing Then
    If headerRight.left <> 0 Or headerRight.top <> 0 Then
      If repair = True Then
        headerRight.top = 0
        headerRight.left = 0
      Else
        issuesObjects.Add headerRight, headerRight.ObjectName
      End If
      count = count + 1
      End If
    
  End If
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper

  
  If uf.ToggleButton1.value = True Then
    If Not bottomRight Is Nothing Then
      Call BringToFront(shoulderRight)
    End If
    If Not bottomRight Is Nothing Then
      Call BringToFront(bottomRight)
    End If
    If Not bottomLeft Is Nothing Then
      Call BringToFront(bottomLeft)
    End If
    If Not bottomLeft Is Nothing Then
      Call BringToFront(titleText)
    End If
    If Not bottomRight Is Nothing Then
      Call BringToFront(shoulderRight)
    End If
    If Not bottomRight Is Nothing Then
      Call BringToFront(bottomRight)
    End If
    If Not bottomLeft Is Nothing Then
      Call BringToFront(bottomLeft)
    End If
    ActiveDocument.Selection.DeselectAll
    count = count + 1
  End If
  
  If repair = False Then
    Dim highlightMain
    Set highlightMain = CreateHighlights(issuesObjects, "Border Alignment") 'will only highlight what was moved already
  End If
  
  HarmonizePopups = count
    
End Function
Sub BringToFront(o As HMIObject)

    ActiveDocument.Selection.DeselectAll
    o.Selected = True
    ActiveDocument.Selection.BringToFront

End Sub
Sub SendToBack(o As HMIObject)

    ActiveDocument.Selection.DeselectAll
    o.Selected = True
    ActiveDocument.Selection.SendToBack

End Sub
Sub solvePopupLayers()

Dim bottomLeft As HMIObject
Dim bottomRight As HMIObject
Dim shoulderLeft As HMIObject
Dim shoulderRight As HMIObject
Dim titleText As HMIObject
Dim headerRight As HMIObject
Dim headerLeft As HMIObject
Dim objs As HMIObjects

Set objs = ActiveDocument.HMIObjects

  On Error Resume Next
  Set bottomLeft = objs("@BottomLeft")
  On Error Resume Next
  Set bottomRight = objs("@BottomRight")
  On Error Resume Next
  Set shoulderLeft = objs("@ShoulderLeft")
  On Error Resume Next
  Set shoulderRight = objs("@ShoulderRight")
  On Error Resume Next
  Set headerLeft = objs("@HeaderLeft")
  On Error Resume Next
  Set headerRight = objs("@HeaderRight")
  On Error Resume Next
  Set titleText = objs("@TitleText")
  
  If Not bottomRight Is Nothing Then
    Call BringToFront(shoulderRight)
  End If
  If Not bottomRight Is Nothing Then
    Call BringToFront(bottomRight)
  End If
  If Not bottomLeft Is Nothing Then
    Call BringToFront(bottomLeft)
  End If
  If Not bottomLeft Is Nothing Then
    Call BringToFront(titleText)
  End If
  If Not bottomRight Is Nothing Then
    Call BringToFront(shoulderRight)
  End If
  If Not bottomRight Is Nothing Then
    Call BringToFront(bottomRight)
  End If
  If Not bottomLeft Is Nothing Then
    Call BringToFront(bottomLeft)
  End If
  ActiveDocument.Selection.DeselectAll
  
End Sub
Public Function HarmonizeIOBackColors(repair As Boolean) As Integer

Dim o As HMIObject
Dim oFrame As HMIObject
Dim objs As HMIObjects
Set objs = ActiveDocument.HMIObjects
Dim issuesObjects As New Collection
Dim count As Integer
Dim prop As HMIProperty
Dim props As HMIProperties

Dim oActRef As New Collection 'customized objects
Set oActRef = joinCollections(GetHMIObjectsByType("HMICustomizedObject", objs, "ActV"), GetHMIObjectsByType("HMICustomizedObject", objs, "RefV"))
Set oActRef = joinCollections(oActRef, GetHMIObjectsByType("HMICustomizedObject", objs, "AnaMsr"))
Set oActRef = joinCollections(oActRef, GetHMIObjectsByType("HMICustomizedObject", objs, "_Pid"))

Dim oPoly As New Collection
Set oPoly = GetHMIObjectsByType("HMIPolygon", objs)

For Each o In oActRef
  If Not o.Properties("BackColor") Is Nothing Then
    'Debug.Print o.Properties("BackColor").value & ", " & ActiveDocument.BackColor
    If o.Properties("BackColor").value <> ActiveDocument.BackColor.value Then
        issuesObjects.Add o, o.ObjectName
        
        Set oFrame = GetObjectFromCollectionByCoordinates(oPoly, o.left, o.top, 3, 3)
        On Error Resume Next
        If oFrame.Properties("BackColor").value <> ActiveDocument.BackColor.value Then
            issuesObjects.Add oFrame, oFrame.ObjectName
            End If
      End If
  End If
      
    Set props = o.Properties
    For Each prop In props
      If InStr(1, UCase(prop.DisplayName), UCase("Background Color"), vbBinaryCompare) Then
          If prop.value <> ActiveDocument.BackColor.value Then
          On Error Resume Next
            issuesObjects.Add o, o.ObjectName
            
            Set oFrame = GetObjectFromCollectionByCoordinates(oPoly, o.left, o.top, 3, 3)
            On Error Resume Next
            If oFrame.Properties("BackColor").value <> ActiveDocument.BackColor.value Then
                issuesObjects.Add oFrame, oFrame.ObjectName
                End If
            End If
        End If
    Next
  
Next

If repair = False Then
  Dim highlightMain
  Set highlightMain = CreateHighlights(issuesObjects, "IO BackColor")
Else
  For Each o In issuesObjects
  
    o.Properties("BackColor").value = ActiveDocument.BackColor.value
    Set props = o.Properties
    
    For Each prop In props
      If InStr(1, UCase(prop.DisplayName), UCase("Background Color"), vbBinaryCompare) Or InStr(1, UCase(prop.name), UCase("BackColor"), vbBinaryCompare) Then
          prop.value = ActiveDocument.BackColor.value
        End If
    Next
    
    If InStr(1, o.ObjectName, "DatCls2", vbBinaryCompare) = 0 And ActiveDocument.BackColor.value = -2147483556 Then
      o.ObjectName = Replace(o.ObjectName, "DatCls3", "DatCls2", , , vbBinaryCompare)
    ElseIf InStr(1, o.ObjectName, "DatCls3", vbBinaryCompare) = 0 And ActiveDocument.BackColor.value = -2147483557 Then
      o.ObjectName = Replace(o.ObjectName, "DatCls2", "DatCls3", , , vbBinaryCompare)
    End If
    
  Next
End If

count = issuesObjects.count

  HarmonizeIOBackColors = count
  
End Function

Public Function HarmonizeUnits(repair As Boolean) As Integer

  Dim reasons As String

  Dim ioFields As New Collection
  Dim io As HMIObject
  Dim startTime, endTime
  
  issueCount = 0
  
  Rem restrict searches to relevant objects
  Dim objs As HMIObjects
  Set objs = ActiveDocument.HMIObjects
  
  startTime = Timer()
  
  Dim oActObjLeft As HMIObject
  Dim oActObjRight As HMIObject
  
  Dim oActRef As New Collection 'customized objects
  Set oActRef = joinCollections(GetHMIObjectsByType("HMICustomizedObject", objs, "ActV"), GetHMIObjectsByType("HMICustomizedObject", objs, "RefV"))
  Set oActRef = joinCollections(oActRef, GetHMIObjectsByType("HMICustomizedObject", objs, "AnaMsr"))
  Set oActRef = joinCollections(oActRef, GetHMIObjectsByType("HMICustomizedObject", objs, "_Pid"))
  
  Dim ounitSet As New Collection
  Set ounitSet = joinCollections(GetHMIObjectsByType("HMIIOField", objs, "@V3_SMS_Unit"), GetHMIObjectsByType("HMIIOField", objs, "Unit_Met"))
  
  Dim objsInContainer As New Collection
  Dim oconts As New Collection
      
  Dim ioLeft
  Dim ioTop
  
  Dim isPolygonContainer
  
  Dim issuesIo As New Collection
    
  Dim potentialContainers As New Collection
  Set potentialContainers = GetHMIObjectsByType("HMIRectangle", objs, , 5000, -2147483557)
      
  Set ioFields = GetHMIObjectsByType("HMIIOField", ActiveDocument.HMIObjects)
  
  For Each io In ounitSet
  ActiveDocument.Selection.DeselectAll
    'check if object is near actV
    'check if object has parantheses or not for appropriate cases
    
    'If io.FONTBOLD = True Then GoTo nextio
    'io.Selected = True
    
    'find if in container !!!!!!!!!!!!!!!!!!!!1111
    Set oconts = Nothing
    Set oconts = getContainers(io, objs, 3)
    If oconts.count = 1 Then
      Set objsInContainer = ObjectsInContainer(oconts(1), objs, 5)
      If oconts(1).Type = "HMIPolygon" Then
        isPolygonContainer = 1
      ElseIf oconts(1).Type = "HMIRectangle" Then
        isPolygonContainer = 0
      End If
      'Call selectCollection(objsInContainer)
    ElseIf oconts.count > 1 Then
      If oconts(1).left = oconts(2).left And oconts(1).top = oconts(2).top And oconts(1).width = oconts(2).width And oconts(1).height = oconts(2).height Then
        objs(oconts(2).ObjectName).Delete 'delete duplicate
        oconts.Remove (2)
      Else
        'MsgBox "more than 1 container for this io field"
      End If
    'Else: MsgBox "page has no container": Exit For
    End If
    'Call selectCollection(oconts)
    
    'find object
    Set oActObjLeft = GetObjectFromCollectionByCoordinates(oActRef, io.left - 70, io.top - 5, 15, 10)
    Set oActObjRight = GetObjectFromCollectionByCoordinates(oActRef, io.left + io.width, io.top - 5, 10, 10)
    
    'check for both
    If Not oActObjRight Is Nothing And Not oActObjLeft Is Nothing Then
      Debug.Print "getting both right and left related outputs here here"
      If io.Type = "HMIIOField" Then
        If InStr(1, io.InputValue, "[", vbBinaryCompare) Then
          Set oActObjLeft = Nothing
          Else
          Set oActObjRight = Nothing
          End If
      ElseIf io.Type = "HMIStaticText" Then
        If InStr(1, io.text, "[", vbBinaryCompare) Then
          Set oActObjLeft = Nothing
          Else
          Set oActObjRight = Nothing
          End If
      End If
    End If
    
    If Not oActObjLeft Is Nothing And isPolygonContainer = 0 Then
      'oActObjLeft.Selected = True
        If oActObjLeft.width = 78 Then 'has fph handling
        ioLeft = CInt(oActObjLeft.left + oActObjLeft.width + 1) 'account if width is different?
      Else
        ioLeft = CInt(oActObjLeft.left + oActObjLeft.width + 1 + 3) 'account if width is different?
      End If
      ioTop = oActObjLeft.top + RoundUp((oActObjLeft.height - io.height) / 2) - 1
'      If repair = False Then
        If io.left <> ioLeft Or io.top <> ioTop Or io.AdaptBorder <> True Or io.AlignmentLeft <> 0 Or io.AlignmentTop <> 1 _
        Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
        Or InStr(1, io.InputValue, "[", vbBinaryCompare) _
        Or InStr(1, io.InputValue, "]", vbBinaryCompare) Then
          issuesIo.Add io, io.ObjectName
          
          If io.left <> ioLeft Then reasons = reasons & vbCrLf & "left"
          If io.top <> ioTop Then reasons = reasons & vbCrLf & "top"
          If io.AdaptBorder <> True Then reasons = reasons & vbCrLf & "adaptborder not true"
          If io.AlignmentLeft <> 0 Then reasons = reasons & vbCrLf & "alignmentleft not 0"
          If io.AlignmentTop <> 1 Then reasons = reasons & vbCrLf & "alignmenttop not 1"
          If InStr(1, io.InputValue, " ", vbBinaryCompare) Then reasons = reasons & vbCrLf & "has space in input value"
          If InStr(1, io.InputValue, "[", vbBinaryCompare) Then reasons = reasons & vbCrLf & "has [ in input value"
          If InStr(1, io.InputValue, "]", vbBinaryCompare) Then reasons = reasons & vbCrLf & "has ] in input value"
          
        End If
'      End If
      If repair = True Then
        io.AdaptBorder = True
        io.left = ioLeft
        io.top = ioTop
        io.AlignmentLeft = 0
        io.AlignmentTop = 1
        io.FONTBOLD = False
        If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
        If InStr(1, io.InputValue, "[", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "[", "", , , vbBinaryCompare)
        If InStr(1, io.InputValue, "]", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "]", "", , , vbBinaryCompare)
      End If
    ElseIf Not oActObjRight Is Nothing And isPolygonContainer = 0 Then
      'oActObjRight.Selected = True
      If oActObjRight.width = 78 Then 'has fph handling
        ioLeft = CInt(oActObjRight.left - io.width) 'account if width is different?
      Else
        ioLeft = CInt(oActObjRight.left - io.width - 3) 'account if width is different?
      End If
      ioTop = oActObjRight.top + RoundUp((oActObjRight.height - io.height) / 2) - 1
      'ActiveDocument.selection.DeselectAll
      'io.Selected = True
      'oActObjRight.Selected = True
'      If repair = False Then
        If io.left <> ioLeft Or io.top <> ioTop Or io.AlignmentLeft <> 2 Or io.AlignmentTop <> 1 _
        Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
        Or InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 _
        Or InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then
          issuesIo.Add io, io.ObjectName
        End If
'      End If
      If repair = True Then
        io.left = ioLeft
        io.top = ioTop
        io.AlignmentLeft = 2
        io.AlignmentTop = 1
        io.FONTBOLD = False
        If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
        If InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 Then io.InputValue = "[" & io.InputValue
        If InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then io.InputValue = io.InputValue & "]"
      End If
    ElseIf isPolygonContainer = 1 Then
      'do different actions ...should be bolded, et al
      If Not oActObjLeft Is Nothing Then
        If oActObjLeft.width = 78 Then 'has fph handling
          ioLeft = CInt(oActObjLeft.left + oActObjLeft.width - 3)
        Else
          ioLeft = CInt(oActObjLeft.left + oActObjLeft.width)
        End If
        ioTop = oActObjLeft.top + RoundUp((oActObjLeft.height - io.height) / 2)
          If io.left <> ioLeft Or io.top <> ioTop Or io.AdaptBorder <> True Or io.AlignmentLeft <> 0 Or io.AlignmentTop <> 1 _
          Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
          Or InStr(1, io.InputValue, "[", vbBinaryCompare) _
          Or InStr(1, io.InputValue, "]", vbBinaryCompare) Then
            issuesIo.Add io, io.ObjectName
          End If
        If repair = True Then
          io.AdaptBorder = True
          io.left = ioLeft
          io.top = ioTop
          io.AlignmentLeft = 0
          io.AlignmentTop = 1
          io.FONTBOLD = True
          If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
          If InStr(1, io.InputValue, "[", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "[", "", , , vbBinaryCompare)
          If InStr(1, io.InputValue, "]", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, "]", "", , , vbBinaryCompare)
        End If
      ElseIf Not oActObjRight Is Nothing Then
        If oActObjRight.width = 78 Then 'has fph handling
          ioLeft = CInt(oActObjRight.left - io.width) 'account if width is different?
        Else
          ioLeft = CInt(oActObjRight.left - io.width - 3) 'account if width is different?
        End If
        ioTop = oActObjRight.top + RoundUp((oActObjRight.height - io.height) / 2) - 1
          If io.left <> ioLeft Or io.top <> ioTop Or io.AlignmentLeft <> 2 Or io.AlignmentTop <> 1 _
          Or InStr(1, io.InputValue, " ", vbBinaryCompare) _
          Or InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 _
          Or InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then
            issuesIo.Add io, io.ObjectName
          End If
        If repair = True Then
          io.left = ioLeft
          io.top = ioTop
          io.AlignmentLeft = 2
          io.AlignmentTop = 1
          io.FONTBOLD = False
          If InStr(1, io.InputValue, " ", vbBinaryCompare) > 0 Then io.InputValue = Replace(io.InputValue, " ", "", , , vbBinaryCompare)
          If InStr(1, io.InputValue, "[", vbBinaryCompare) = 0 Then io.InputValue = "[" & io.InputValue
          If InStr(1, io.InputValue, "]", vbBinaryCompare) = 0 Then io.InputValue = io.InputValue & "]"
        End If
      End If
    Else
      'MsgBox "no ActObj !"
    End If
nextio:
  Next
  
  Set oconts = Nothing
  If repair = False Then Set oconts = CreateHighlights(issuesIo)  'reuse coll for useless highlights...
  
  Dim what As String
  If repair = True Then
    what = "repair of unit io fields"
  Else
    what = "check of unit io fields"
  End If
  
  HarmonizeUnits = issuesIo.count
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  uf.ListBox1.AddItem ("Ran " & what & ", " & issuesIo.count & " issues in " & FormatNumber(Timer() - startTime, 3) & " seconds")
End Function

Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem //////////////////////////// FUNCTIONS //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Rem /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Rem *************************** Buttons *****************************************************************************************************************************************************

Public Function GetObjectFromCollectionByCoordinates(col As Collection, l As Variant, t As Variant, toleranceLeft As Integer, toleranceTop As Integer) As HMIObject
  Dim elem As HMIObject
  Dim o As HMIObject
  
  For Each elem In col
    If Abs(elem.left - l) < toleranceLeft And Abs(elem.top - t) < toleranceTop Then
      Set o = elem
      Exit For
    End If
  Next
  
  Set GetObjectFromCollectionByCoordinates = o
End Function

Public Function GetContainedObjects(col As HMICollection, o As HMIObject) As Collection
  Dim elem As HMIObject
  
  Dim hmiCol As New Collection
  Dim l
  Dim t
  Dim toleranceLeft
  Dim toleranceTop
  
  l = o.left + o.width / 2
  t = o.top + o.height / 2
  toleranceLeft = o.width / 2 - 1
  toleranceTop = o.height / 2 + 3
  
  For Each elem In col
    If Abs(elem.left - l) < toleranceLeft And Abs(elem.top - t) < toleranceTop Then
      hmiCol.Add elem, elem.ObjectName
'      elem.Selected = True
'      Debug.Print "0"
    End If
  Next
  
  Set GetContainedObjects = hmiCol
End Function

Public Function GetNearbyVerticalButtonsGroup(c As Variant, buttons As Collection, heightMargin As Integer, widthMargin As Integer, leftMargin As Integer) As Collection
  Dim b As HMIButton
  Dim cl1 As New Collection
  Dim cl2 As New Collection
  Dim elem As HMIButton
  Dim col As New Collection

  Rem go on a search from the found button, upwards until interval gets too large then do not include
  Rem same goes for search to bottom, and remove from main button collection set from which to get related cols

  c.Selected = True
  On Error Resume Next
  buttons.Remove c.ObjectName
  For Each b In buttons
    Rem less than 21 deviation from left/right
    Rem spacing margin tolerance
    Rem buttons will be grouped as long as there's less than 5 pixels difference in width
    b.Selected = True
    If Abs(b.left - c.left) < leftMargin And (Abs(c.top - b.top - b.height) < heightMargin Or Abs(c.top + c.height - b.top) < heightMargin) And Abs(b.width - c.width) < widthMargin Then
      ActiveDocument.Selection.DeselectAll
      b.Selected = True
      buttons.Remove b.ObjectName
      Set cl2 = GetNearbyVerticalButtonsGroup(b, buttons, heightMargin, widthMargin, leftMargin)
    End If
  Next

  For Each elem In ActiveDocument.Selection
    col.Add elem, elem.ObjectName
  Next

  Set GetNearbyVerticalButtonsGroup = col
End Function

Public Function IndexOf(ByVal coll As Collection, ByVal Item As Variant) As Long
    Dim i As Long
    For i = 1 To coll.count
        If coll(i).ObjectName.value = Item.ObjectName.value Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

Public Function getButtonGroups(col As Collection) As Collection
  Dim i As Integer
  Dim j As Long
  Dim l As Integer
  
  Dim indexLeft
  Dim spce
  
  Dim elem As HMIObject
  
  Dim colSpaces As New Collection
  Dim colMain As New Collection
  Dim colVert As New Collection
  Dim spcSec As New Collection
  Dim emptyCol As New Collection
  Dim emptyColSpc As New Collection
  Dim sorted As New Collection
  
  Dim colSec As New Collection
  Dim highlightMain As New Collection
  Dim highlightSec As New Collection
  Dim difLeft As New Collection
  Dim minLeft As Integer
  
  Set sorted = SortCollectionByProperty(col, "Left")
  
  indexLeft = 0
  i = 1
  colVert.Add emptyCol, CStr(i)
  For Each elem In sorted
    If indexLeft = 0 Then indexLeft = elem.left
    If elem.left > indexLeft And Abs(elem.left - indexLeft) > 10 Then  '10 is the precision for error along vertical group, even 50 should be fine
      i = i + 1
      Set emptyCol = Nothing
      colVert.Add emptyCol
      indexLeft = elem.left
    End If
    colVert(i).Add elem
  Next
  
  For j = 1 To colVert.count
    i = 1: l = 1
    Set colSec = Nothing: Set spcSec = Nothing
    Set emptyCol = Nothing:  Set emptyColSpc = Nothing
    Set sorted = SortCollectionByProperty(colVert(j), "Top")
    colSec.Add emptyCol
    spcSec.Add emptyColSpc
    For Each elem In sorted
      If i > 0 And i < sorted.count Then
        spce = Abs(sorted(i).top + sorted(i).height - sorted(i + 1).top)
        If spce > 20 Then
          colSec(l).Add sorted(i), sorted(i).ObjectName.value
          l = l + 1
          Set emptyCol = Nothing: colSec.Add emptyCol
          Set emptyColSpc = Nothing: spcSec.Add emptyColSpc
        Else
          colSec(l).Add sorted(i), sorted(i).ObjectName.value
          spcSec(l).Add spce, CStr(i)
        End If
        i = i + 1
        If i = sorted.count Then colSec(l).Add sorted(i), sorted(i).ObjectName.value
      End If
    Next
  colMain.Add colSec
  colSpaces.Add spcSec
  Next
  
  Set colSec = Nothing
  colSec.Add colMain
  colSec.Add colSpaces
    
Set getButtonGroups = colSec
End Function

Rem *************************** Containers **************************************************************************************************************************************************

Public Function getContainers(o As Variant, os As HMIObjects, maxDeviationOutside As Integer) As Collection
  Dim sobj As Variant
  Dim prop As HMIProperty
  Dim conts As New Collection
  Dim potentialContainers As Collection
  Dim val1 As Long
  Dim val2 As Long
  
  If InStr(1, ActiveDocument.name, "Class3", vbBinaryCompare) = 0 And ActiveDocument.BackColor.value = -2147483556 Then
    val1 = -2147483561 '87
    val2 = -2147483516 '132
  ElseIf InStr(1, ActiveDocument.name, "Class3", vbBinaryCompare) > 0 And ActiveDocument.BackColor.value = -2147483557 Then
    val1 = -2147483527 '121
    val2 = -2147483516 '132
  End If
  
  Set potentialContainers = GetHMIObjectsByType("HMIRectangle", os, , 5000, ActiveDocument.BackColor.value, val1)
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIPolygon", os, , 10, ActiveDocument.BackColor.value, val1))
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIPolygon", os, , 10, ActiveDocument.BackColor.value, val1))
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIRectangle", os, , 5000, ActiveDocument.BackColor.value, val1))
  Set potentialContainers = joinCollections(potentialContainers, GetHMIObjectsByType("HMIRectangle", os, , 5000, ActiveDocument.BackColor.value, val2))
  
  'Call selectCollection(potentialContainers)
    
  For Each sobj In potentialContainers
  Set prop = sobj.BackColor
    If ((o.left + o.width) - (sobj.left + sobj.width)) <= maxDeviationOutside And (o.left - sobj.left) >= -maxDeviationOutside And _
    ((o.top + o.height) - (sobj.top + sobj.height)) <= maxDeviationOutside And (o.top - sobj.top) >= -maxDeviationOutside And _
    sobj.Type <> "HMIGroup" Then
      If InStr(1, sobj.ObjectName, "@", vbTextCompare) = 0 Then 'layer < 29 ?
        conts.Add sobj
      End If
    End If
  Next
  
  'Call selectCollection(conts)
  
  Set getContainers = conts
End Function

Private Sub ArrangeContainerTitle(t As HMIStaticText, c As HMIRectangle, off As Integer)
  Dim count As Integer
  
  If left(t.text, 1) <> " " Then t.text = " " & t.text
  If right(t.text, 1) <> " " Then t.text = t.text & " "
  t.left = c.left + 10 - off
  count = (CountChrInString(t.text, vbCrLf) + 1) * 2
    t.top = c.top - t.height / count
End Sub

Private Sub ArrangeContainerTitleUnit(u As HMIObject, t As HMIStaticText, c As HMIRectangle)
  'If CInt(ounit.left) <> CInt(otitle.left + otitle.Width - 1) Or Right(ounit.InputValue, 1) <> " " Or ounit.top <> otitle.top Or ounit.AlignmentLeft <> 0 Then
  
  Dim vb As HMIScriptInfo
  Dim src As String
    
  If u.Type = "HMIIOField" Then
    u.InputValue = Replace(u.InputValue, " ", "", , , vbBinaryCompare)
    
    If left(u.InputValue, 1) = " " Then u.InputValue = right(u.InputValue, Len(u.InputValue - 1))
    If right(u.InputValue, 1) <> " " Then u.InputValue = u.InputValue & " "
    
    u.left = t.left + t.width - 2
    u.AdaptBorder = True
    u.AlignmentLeft = 0
    u.top = t.top
    Set vb = u.OutputValue.Dynamic
      With vb 'will have to check sourcecode when selecting units to change...
        src = .sourceCode
        If InStr(1, src, """]""", vbBinaryCompare) Then
          .sourceCode = Replace(.sourceCode, """]""", """] """, , , vbBinaryCompare)
        End If
      End With
  ElseIf u.Type = "HMIStaticText" Then
    u.FONTBOLD = True
    u.left = t.left + t.width - 2
    u.AdaptBorder = True
    u.AlignmentLeft = 0
    u.top = t.top
    
    On Error Resume Next
    Set vb = u.text.Dynamic
    
    If Not vb Is Nothing Then
      With vb 'will have to check sourcecode when selecting units to change...
        src = .sourceCode
        If InStr(1, src, """]""", vbBinaryCompare) Then
          .sourceCode = Replace(.sourceCode, """]""", """] """, , , vbBinaryCompare)
        End If
      End With
    End If
    
    If right(u.text, 1) = "]" Then
      u.text = u.text & " "
    ElseIf right(u.text, 2) = "] " Then
      
    
    End If
    If InStr(1, u.text, "[", vbBinaryCompare) Then
      If left(u.text, 2) = " [" Then
        u.text = Replace(u.text, " [", "[", , , vbBinaryCompare)
      End If
    Else
      u.text = "[" & u.text
    End If
    
    
  End If
    
End Sub

Private Function ObjectsInContainer(c As HMIObject, os As HMIObjects, tol As Integer) As Collection
  Dim o As HMIObject
  Dim col As New Collection
  
  For Each o In os
    If (o.left > c.left Or Abs(o.left - c.left) <= tol) _
    And ((o.left + o.width) <= (c.left + c.width)) _
    And ((o.top < c.top) Or Abs(o.top - c.top) <= tol) _
    And Abs((o.top + o.height) - (c.top + o.height)) <= tol Then
      col.Add o, o.ObjectName
    End If
  Next
  col.Remove (c.ObjectName)
  Set ObjectsInContainer = col
End Function

Rem *************************** Common *****************************************************************************************************************************************************

Public Function SortCollectionByProperty(colInput As Collection, prop As String) As Collection
  Dim iCounter As Integer
  Dim iCounter2 As Integer
  Dim temp As Variant
  
  For iCounter = 1 To colInput.count - 1
      For iCounter2 = iCounter + 1 To colInput.count
          If colInput(iCounter).Properties(prop) > colInput(iCounter2).Properties(prop) Then
             Set temp = colInput(iCounter2)
             colInput.Remove iCounter2
             colInput.Add temp, temp.ObjectName, iCounter
          End If
      Next iCounter2
  Next iCounter
  Set SortCollectionByProperty = colInput
End Function

Public Function joinCollections(col1 As Collection, col2 As Collection) As Collection
  Dim mainCol As New Collection
  Dim elem As Variant
  
  For Each elem In col1
    On Error Resume Next
    mainCol.Add elem, elem.ObjectName
  Next
  For Each elem In col2
    On Error Resume Next
    mainCol.Add elem, elem.ObjectName
  Next
  
  Set joinCollections = mainCol
End Function

Sub selectCollection(col As Collection)
  Dim elem As HMIObject
  
  For Each elem In col
    elem.Selected = True
    Debug.Print elem.ObjectName
  Next
End Sub

Private Function CreateLeftHighlights(col As Collection, leftValue As Integer) As Collection
  Dim a As HMIObject
  Dim newCol As New Collection
  Dim o As HMILine
  Dim l, t, w, h

  For Each a In col
    l = a.left 'CInt(s(0))
    t = a.top 'CInt(s(1))
    w = a.width 'CInt(s(2))
    h = a.height 'CInt(s(3))
    Set o = ActiveDocument.HMIObjects.AddHMIObject("HarmonizationAux", "HMILine")
      With o
        .left = leftValue
        .top = a.top
        .height = a.height
        .width = 0
        .BorderColor = RGB(255, 201, 14) 'gold
        .Transparency = 50
        .BorderStyle = 512
        .BorderEndStyle = 514
        .BorderBackColor = RGB(255, 201, 14) 'gold
        .BorderWidth = 6
        .GlobalColorScheme = False
        .GlobalShadow = False
        .left = .left - (.BorderWidth / 2) - 1
      End With
    newCol.Add o
  Next

  Set CreateLeftHighlights = newCol
End Function

Private Function CreateHighlights(col As Collection, Optional str As String) As Collection
  Dim a As HMIObject
  Dim newCol As New Collection
  Dim o As HMIStaticText
  Dim l, t, w, h

  For Each a In col
    's = Split(a, ",", , vbBinaryCompare)
    l = a.left 'CInt(s(0))
    t = a.top 'CInt(s(1))
    w = a.width 'CInt(s(2))
    h = a.height 'CInt(s(3))
    Set o = ActiveDocument.HMIObjects.AddHMIObject("HarmonizationAux", "HMIStaticText")
    o.left = l
    o.top = t
    o.width = w
    o.height = h
    o.Transparency = 50
    o.BackColor = RGB(255, 201, 14) 'gold
    o.Layer = 27
    o.GlobalColorScheme = False
    o.GlobalShadow = False
    o.DrawInsideFrame = False
    
    If IsMissing(str) = False Then
      o.text = str
      o.FONTSIZE = 11
      o.FONTBOLD = True
      o.AlignmentLeft = 0
      o.AlignmentTop = 0
    End If
    
    newCol.Add o
  Next

  Set CreateHighlights = newCol
End Function

Sub DeleteObjectsFoundInCol(col As Collection)
  Dim elem As HMIObject
  For Each elem In col
    elem.Delete
  Next
End Sub

Sub DeleteHarmonizationAuxiliaries()
  Dim obj As HMIObject
  Dim startt
  startt = Timer()
  Dim objs As HMIObjects
  Set objs = ActiveDocument.HMIObjects
  
  issueCount = 0
  For Each obj In objs
    If InStr(1, obj.ObjectName, "Harmonization", vbBinaryCompare) = 1 Then
      obj.Delete
      issueCount = issueCount + 1
    End If
  Next
  
  Dim uf As UserForm
  Set uf = HarmonizationHelper
  uf.ListBox1.AddItem ("Cleared issue " & issueCount & " highlights in " & FormatNumber(Timer() - startt, 3) & " seconds")
End Sub

Public Function GetHMIObjectsByType(typ As String, os As HMIObjects, Optional str As String, Optional area As Variant, Optional bacColor As Long, Optional borColor As Long, Optional textStartsWith As String, Optional textEndsWith As String) As Collection
' restrict to certain type
' added area to check for area occupied by object ..
  Dim o As HMIObject
  Dim ba
  Dim bo
  Dim col As New Collection
  Dim elem As HMIObject
  
  Dim colSearchResults As HMICollection
  Dim objMember As HMIObject
  Dim iResult As Integer
  Dim strName As String
  
  Set colSearchResults = os.Find(typ)

  For Each o In colSearchResults
  
    If o.Type = typ Then
      col.Add o, o.ObjectName
    End If
    
    If IsMissing(str) = False Then
      If str <> "" Then
        If InStr(1, o.ObjectName, str, vbBinaryCompare) Then
        Else
          On Error Resume Next
          col.Remove (o.ObjectName)
        End If
      End If
    End If
    
    If IsMissing(area) = False Then
      If o.width * o.height >= area Then
      Else
        On Error Resume Next
        col.Remove (o.ObjectName)
      End If
    End If
    
    If IsMissing(bacColor) = False And bacColor <> 0 Then
      On Error Resume Next
      ba = o.BackColor.value
      If ba <> 0 Then
        If ba = bacColor Then 'Or ba = bacColor - 1 Or ba = bacColor + 1 Then
        Else
          On Error Resume Next
          col.Remove (o.ObjectName)
        End If
      End If
    End If
    
    If IsMissing(borColor) = False And borColor <> 0 Then
      On Error Resume Next
      bo = o.BorderColor.value
      If bo <> 0 Then
        If bo = borColor Then 'Or bo = borColor - 1 Or bo = borColor + 1 Then
        Else
          On Error Resume Next
          col.Remove (o.ObjectName)
        End If
      End If
    End If
    
        Dim prop As HMIProperty
        
    If IsMissing(textStartsWith) = False Then
      If textStartsWith <> "" Then
      
        On Error Resume Next
        Set prop = o.text
        
        If Not prop Is Nothing Then
          If InStr(1, left(o.text, 3), textStartsWith, vbBinaryCompare) Then 'Or bo = borColor - 1 Or bo = borColor + 1 Then
          Else
            On Error Resume Next
            col.Remove (o.ObjectName)
          End If
        End If
        
      End If
    End If
    
    If IsMissing(textEndsWith) = False Then
      If textEndsWith <> "" Then
      
        On Error Resume Next
        Set prop = o.text
        
        If Not prop Is Nothing Then
          If InStr(1, right(o.text, 3), textEndsWith, vbBinaryCompare) Then 'Or bo = borColor - 1 Or bo = borColor + 1 Then
          Else
            On Error Resume Next
            col.Remove (o.ObjectName)
          End If
        End If
        
      End If
    End If
    
    
  Next
  
  Set GetHMIObjectsByType = col
End Function

Public Function InCollection(col As Collection, Key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = col(Key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If
End Function

Function RoundUp(ByVal d As Double) As Integer
    Dim result As Integer
    result = Math.Round(d)
    If result >= d Then
        RoundUp = result
    Else
        RoundUp = result + 1
    End If
End Function
Sub teeeest()
Dim o As HMIStaticText
Set o = ActiveDocument.HMIObjects("Static Text41")
Debug.Print o.ForeColor.value
End Sub
Public Function CheckMarginsInsideContainers(repair As Boolean) As Integer

  Dim myCol As New Collection
  
  Dim o As HMIObject

  Dim col As HMICollection
  Set col = ActiveDocument.HMIObjects.Find(PropertyName:="Left")
  
  Dim cont As New Collection
  Set cont = GetAllContainers()
  
  Dim issues As New Collection
  
  Dim c As HMIObject
  
  Dim reasons As String
  
  Dim prop As HMIProperty
    
  For Each c In cont
    ActiveDocument.Selection.DeselectAll
    c.Selected = True
  
    Set myCol = GetContainedObjects(col, c)
    
    For Each o In myCol
      reasons = ""
      
      If o.left < c.left + 10 Then 'too little left margin
        On Error Resume Next
        issues.Add o, o.ObjectName
'        o.Selected = True
        reasons = reasons & ", left margin not 10px"
      End If
      
      If (o.left + o.width) > (c.left + c.width) - 10 Then 'too little right margin
        On Error Resume Next
        issues.Add o, o.ObjectName
        o.Selected = True
        reasons = reasons & ", right margin not 10px"
        
        If repair = True Then
          On Error Resume Next
          Set prop = o.ForeColor
          If Not prop Is Nothing Then
            If o.ForeColor = -2147483640 Then
            o.Selected = True
              o.left = c.left + c.width - 10 - o.width
            End If
          End If
        End If
        
      End If
      
      If o.top < c.top + 15 Then 'too little top margin
        On Error Resume Next
        issues.Add o, o.ObjectName
'        o.Selected = True
        reasons = reasons & ", top margin not 15px"
        
        If repair = True Then
          On Error Resume Next
          Set prop = o.ForeColor
          If Not prop Is Nothing Then
            If o.ForeColor = -2147483640 Then
            o.Selected = True
              o.top = c.top + 15
            End If
          End If
        End If
        
      End If
      
      If o.top + o.height > c.top + c.height - 10 Then 'too little bottom margin
        On Error Resume Next
        issues.Add o, o.ObjectName
'        o.Selected = True
        reasons = reasons & ", bottom margin not 10px"
      End If
      
      If reasons <> "" And repair = False Then
        Dim highlightMain
        Dim workaround As New Collection
        workaround.Add o, o.ObjectName
        Set highlightMain = CreateHighlights(workaround, right(reasons, Len(reasons) - 2))
        Set highlightMain = Nothing
        Set workaround = Nothing
        Debug.Print "test"
      End If

    Next o
    Set myCol = Nothing
  Next c
  
  
CheckMarginsInsideContainers = issues.count

End Function

Public Function GetAllContainers() As Collection

  Dim o As HMIObject
  Dim obj As HMIObject
  Dim oconts As New Collection
  Dim foundContainers As New Collection
  Dim oSMS As New Collection 'customized objects
  Set oSMS = GetHMIObjectsByType("HMICustomizedObject", ActiveDocument.HMIObjects, "@V3")
  
  For Each obj In oSMS 'looks to standard object, then at possibly existing container from potential set
    ActiveDocument.Selection.DeselectAll 'clear selection for next steps
    
    Set oconts = getContainers(obj, ActiveDocument.HMIObjects, 0)
    For Each o In oconts
      'o.Selected = True
      On Error Resume Next
      foundContainers.Add o, o.ObjectName
    Next
  Next
  
  Set GetAllContainers = foundContainers
  
  Debug.Print GetAllContainers.count
  
End Function

Sub FindObjectsByType()

  'VBA272
  Dim colSearchResults As HMICollection
  Dim objMember As HMIObject
  Dim iResult As Integer
  Dim strName As String
  Set colSearchResults = ActiveDocument.HMIObjects.Find(ObjectType:="HMICircle")
  For Each objMember In colSearchResults
    iResult = colSearchResults.count
    strName = objMember.ObjectName
    MsgBox "Found: " & CStr(iResult) & vbCrLf & "Objectname: " & strName
  Next objMember

End Sub

Public Function HarmonizeSequenceButtons(repair As Boolean) As Integer

Dim oPBs As New Collection 'customized objects
Set oPBs = GetHMIObjectsByType("HMIButton", ActiveDocument.HMIObjects, "@V3_SMS_Pb")
Dim col As HMICollection 'customized objects
Set col = ActiveDocument.HMIObjects.Find(ObjectType:="HMICustomizedObject")

Dim myCol As New Collection
Dim o As HMIObject
Dim inner As HMICustomizedObject
Dim udo As HMIUdoObjects
Dim u As HMIObject

Dim top
Dim left
Dim bottom
Dim right

Dim firstObject As HMIObject
Dim secondObject As HMIObject

Dim firstObjectLeft
Dim secondObjectLeft

Dim firstObjectRight
Dim secondObjectRight

Dim topmargin
Dim bottommargin
Dim count As Integer

Dim issues As New Collection

Dim myText As String


Dim highlights As New Collection

For Each o In oPBs
  Set myCol = GetContainedObjects(col, o)
  For Each inner In myCol
  
    Debug.Print inner.ObjectName
    top = 10000
    left = 10000
    bottom = 0
    right = 0
    
    Set udo = inner.HMIUdoObjects
    For Each u In udo
'      myText = ""
'      On Error Resume Next
'      myText = u.text.value
'
'      If myText = "V3" Then
'        GoTo nextU
'        End If
      If InStr(1, u.ObjectName, "Symbol_LimitSafeOp", vbBinaryCompare) Then GoTo nextU
      If InStr(1, u.ObjectName, "Symbol_Triangle", vbBinaryCompare) Then GoTo nextU
      If InStr(1, u.ObjectName, "Symbol_Rectangle", vbBinaryCompare) Then GoTo nextU
      If InStr(1, u.ObjectName, "StatusIcon", vbBinaryCompare) Then GoTo nextU
      If InStr(1, u.ObjectName, "VersionInfo", vbBinaryCompare) Then GoTo nextU
      
      'get smallest values except for V3s
        If top > u.top Then top = u.top
        If left > u.left Then left = u.left
        If bottom < u.top + u.height Then bottom = u.top + u.height
        If right < u.left + u.width Then right = u.left + u.width
        
        Dim shouldbeLeft
        If left - o.left < 10 Then
          shouldbeLeft = 5
          firstObjectLeft = left
          firstObjectRight = right
          Set firstObject = inner
        Else
          secondObjectLeft = left
          secondObjectRight = right
          Set secondObject = inner
        End If
        Debug.Print inner.ObjectName & "," & o.top - top & "," & o.top + o.height - bottom & "," & top & "," & bottom
                
nextU:
    Next u
    
    topmargin = RoundUp((o.height - (bottom - top)) / 2)
    If o.top - top <> -5 Then
      If repair = True Then
        inner.top = inner.top + (o.top - top) + topmargin
      Else
        Dim workaround As New Collection
        workaround.Add inner, inner.ObjectName
        Set highlights = CreateHighlights(workaround, "not entirely centered on the button, should be " & topmargin & " from button.top")
        On Error Resume Next
        issues.Add inner, inner.ObjectName

      End If
      count = count + 1
    End If
    
    Debug.Print "wait"
    
  Next inner
  
  If o.left + 5 <> firstObjectLeft Then
    If repair = True Then
      firstObject.left = o.left + 5 - firstObjectLeft + firstObject.left
    End If
  End If
  If firstObjectRight + 8 <> secondObjectLeft Then
    If repair = True Then
      secondObject.left = firstObjectRight + 8 - secondObjectLeft + secondObject.left
    End If
  End If
  If InStr(1, o.text, "            ", vbBinaryCompare) = 0 Then
    If repair = True Then
      o.text = "            " & LTrim(o.text)
    Else
      On Error Resume Next
      issues.Add o, o.ObjectName
      Dim buttonspacing As New Collection
      buttonspacing.Add o, o.ObjectName
      Set highlights = CreateHighlights(buttonspacing, "leading spaces")
    End If
  End If
    
  
  Debug.Print "test"
Next o
  
HarmonizeSequenceButtons = count

End Function

Sub modifyCustomizedObject()

Dim o As HMICustomizedObject

Set o = ActiveDocument.HMIObjects("@V3_SMS_SignLmp_1")

o.Destroy

Dim s As HMIObject
Dim sel As HMISelectedObjects
Set sel = ActiveDocument.Selection

For Each s In sel
  If InStr(1, s.ObjectName, "VersionInfo", vbTextCompare) Then
    s.Selected = False
    End If
  Next s
  ActiveDocument.Selection.CreateCustomizedObject

End Sub


