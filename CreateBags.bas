Attribute VB_Name = "CreateBags"
Sub CreateBags()

  Dim obj As HMIObject
  Set obj = Nothing
  Dim objsec As HMIObject
  Set objsec = Nothing
  Dim objGroup As HMIGroup
  Dim section As String
  section = "init"
  Rem On Error GoTo ErrorHandler
  Dim bagSubObj As HMIObject
  Dim objs As HMIObjects
  Dim bagSubObjs As HMIGroupedObjects
  Dim numObjs As Integer
  Dim numBagObjs, numBagSubObjs, numObjsFixed As Integer
  Dim objScript, objCScript, objCScriptNew As HMIScriptInfo
  Dim dynamicDialog As HMIDynamicDialog
  Dim dynamicDialogNew As HMIDynamicDialog
  Dim objVarTrigger As HMIVariableTrigger
  Dim objVarTriggerNew As HMIVariableTrigger
  Dim objVarTriggerIONew As HMIVariableTrigger
  Dim objDynDialog As HMIDynamicDialog
  Dim backColor As Long
  Dim sourceCode As String
  Dim d As HMIDynamicCreationType
  Dim action As HMIActionDynamic
  Rem hmiDynamicCreationTypeVariableDirect
  Rem Dim d As HMIDynamicDialog
  Dim dy As Integer
  Dim hmiScript As HMIScriptType
  Dim s As String
  Dim objCircle As HMICircle
  Dim objIOField, objIOField2 As HMIIOField
  Dim convName, bagName As String
  Dim strCode As String
  Dim bagNr As Integer
  Dim bagDensity As Integer
  Dim bagSpac As Double
  Dim bagAngle As Double
  Dim pi As Long
  
  
  
  pi = 3.14159265358979
  bagDensity = 18
  'MsgBox Application.ActiveDocument.Name
  backColorConv = RGB(182, 182, 182)
  backColorBag = RGB(175, 171, 176)
  backColorPec = RGB(0, 0, 84)
  Set objs = Application.ActiveDocument.HMIObjects
  'Set bagSubObjs = ActiveDocument.HMIGroupedObjects
  section = "start"
  
    
For Each obj In objs
If (obj.Layer = 1 And obj.Type = "HMIStaticText") Then
'If obj.ObjectName = "D410" Then
    i = 0
    convName = obj.ObjectName
    If obj.Type = "HMIStaticText" Then
        bagNr = obj.Width / bagDensity '40 cm in pixels is approx 13 px
    End If
    If bagNr < 3 Then
        bagNr = 3
        ElseIf bagNr > 20 Then
        bagNr = 20
        Else
        bagNr = bagNr
    End If
    

    For i = 1 To bagNr

    If i < 10 Then
        bagName = convName & "_BAG_MAP_POS_0" & i
    Else
        bagName = convName & "_BAG_MAP_POS_" & i
    End If

    section = "Any object"

    'create the circle for the bag
    strCode = "long backgroundColour = GetBagStatusColours(lpszObjectName);  return backgroundColour;"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
       objCircle.Left = 0
       objCircle.Top = 0
       objCircle.Radius = 20
       objCircle.GlobalColorScheme = 0
       objCircle.backColor = backColorBag
'       Set objDynDialog = objCircle.Visible.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'BAG_VIS' && """ & bagName & """ > 0")
'            With objDynDialog
'                .ResultType = hmiResultTypeBool
'                .BinaryResultInfo.NegativeValue = 0
'                .BinaryResultInfo.PositiveValue = 1
'            End With
'        Set objCScript = objCircle.backColor.CreateDynamic(hmiDynamicCreationTypeCScript)
'            With objCScript
'            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
'                .sourceCode = "long backgroundColour = GetBagStatusColours(""" & bagName & """);  return backgroundColour;"
'            End With

    'create the upper output field
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField", "HMIIOField")
        objIOField.Left = objCircle.Left + 8
        objIOField.Top = objCircle.Top + 6
        objIOField.Width = 26
        objIOField.Height = 14
        objIOField.FONTNAME = Arial
        objIOField.BoxType = Output
        objIOField.OutputFormat = "099"
        objIOField.LimitMin = -1.79769313486231E+308
        objIOField.LimitMax = 1.79769313486231E+308
        objIOField.backColor = RGB(175, 171, 176)
        objIOField.FillStyle = 65536 'transparent
        objIOField.BorderWidth = 0 'borderweight = 0
        objIOField.FONTBOLD = 1 'font bold
        objIOField.FONTSIZE = 15 'font size 11
        objIOField.AlignmentLeft = 1
        objIOField.AlignmentTop = 1
        objIOField.GlobalColorScheme = 0 'global color  scheme no
        Set objCScript = objIOField.OutputValue.CreateDynamic(hmiDynamicCreationTypeCScript)
            With objCScript
            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
                .sourceCode = "return(((unsigned long) GetTagDouble(""" & bagName & """) % 100000) / 1000);"
            End With

    'create the upper output field
    Set objIOField2 = ActiveDocument.HMIObjects.AddHMIObject("IOField", "HMIIOField")
        objIOField2.Left = objCircle.Left + 8
        objIOField2.Top = objCircle.Top + 20
        objIOField2.Width = 26
        objIOField2.Height = 14
        objIOField2.FONTNAME = Arial
        objIOField2.BoxType = Output
        objIOField2.OutputFormat = "0999"
        objIOField2.LimitMin = -1.79769313486231E+308
        objIOField2.LimitMax = 1.79769313486231E+308
        objIOField2.backColor = RGB(175, 171, 176)
        objIOField2.FillStyle = 65536 'transparent
        objIOField2.BorderWidth = 0 'borderweight = 0
        objIOField2.FONTBOLD = 1 'font bold
        objIOField2.FONTSIZE = 15 'font size 11
        objIOField2.AlignmentLeft = 1
        objIOField2.AlignmentTop = 1
        objIOField2.GlobalColorScheme = 0 'global color  scheme no
        Set objCScript = objIOField2.OutputValue.CreateDynamic(hmiDynamicCreationTypeCScript)
            With objCScript
            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
                .sourceCode = "return((unsigned long) GetTagDouble(""" & bagName & """) % 1000);"
            End With

    'group the created items
    With objCircle
        .Selected = True
    End With
    With objIOField
        .Selected = True
    End With
    With objIOField2
        .Selected = True
    End With


    Set objGroup = ActiveDocument.Selection.CreateGroup
        objGroup.ObjectName = bagName
        objGroup.Layer = 3
        objGroup.Top = obj.Top
        objGroup.Left = obj.Left
        objGroup.Visible = False
        Set objDynDialog = objGroup.Visible.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'BAG_VIS' && '" & bagName & "' > 0")
            With objDynDialog
                .ResultType = hmiResultTypeBool
                .BinaryResultInfo.NegativeValue = 0
                .BinaryResultInfo.PositiveValue = 1
                .Trigger.VariableTriggers.Item(1).CycleType = hmiVariableCycleTypeUponChange
                .Trigger.VariableTriggers.Item(2).CycleType = hmiVariableCycleTypeUponChange
            End With
        Set objCScript = objCircle.backColor.CreateDynamic(hmiDynamicCreationTypeCScript)
            With objCScript
            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
                .sourceCode = "long backgroundColour = GetBagStatusColours(""" & bagName & """);  return backgroundColour;"
            End With
        


        bagSpac = ((obj.Width - objGroup.Width)) / (bagNr - 1)
        If obj.RotationAngle = 0 Then
'            obj.Height = 42
                objGroup.Top = obj.Top
                If Right(objGroup.ObjectName, 2) = "0" & bagNr Then
                    objGroup.Left = obj.Left + obj.Width - objGroup.Width
                ElseIf Right(objGroup.ObjectName, 2) = bagNr Then
                    objGroup.Left = obj.Left + obj.Width - objGroup.Width
                Else
                    objGroup.Left = obj.Left + bagSpac * (i - 1)
                End If
        ElseIf obj.RotationAngle = 360 Then
'            obj.Height = 42
                objGroup.Top = obj.Top
                If Right(objGroup.ObjectName, 2) = "0" & bagNr Then
                    objGroup.Left = obj.Left + obj.Width - objGroup.Width
                ElseIf Right(objGroup.ObjectName, 2) = bagNr Then
                    objGroup.Left = obj.Left + obj.Width - objGroup.Width
                Else
                    objGroup.Left = obj.Left + bagSpac * (i - 1)
                End If
        ElseIf obj.RotationAngle = 270 Then
'            obj.Width = 42
                objGroup.Left = obj.Left + objGroup.Width
                objGroup.Top = obj.Top + bagSpac * (i - 1) - objGroup.Height
        ElseIf obj.RotationAngle = 90 Then
'            obj.Width = 42
                objGroup.Left = obj.Left + objGroup.Width
                objGroup.Top = obj.Top + bagSpac * (i - 1) - objGroup.Height
        End If
'

'        For Each objsec In objs
'        If objsec.Layer = 1 Or objsec.Layer = 2 Then
'
'        If obj.Type = "HMIStaticText" Then
'        If (Left(objsec.ObjectName, 2) = Left(obj.ObjectName, 2)) Then
'        If obj.RotationAngle = 360 Then
'            If ((objsec.Left < obj.Left) And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2)) = -1) And (Abs(objsec.Left - obj.Left) < 300) And (Abs(objsec.Top - obj.Top) < 300) Or (objsec.Left > obj.Left And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2) = 1)) Then
'                'for bags going left-right
'                objGroup.Top = obj.Top
'                If Right(objGroup.ObjectName, 2) = "0" & bagNr Then
'                    objGroup.Left = obj.Left + obj.Width - objGroup.Width
'                ElseIf Right(objGroup.ObjectName, 2) = bagNr Then
'                    objGroup.Left = obj.Left + obj.Width - objGroup.Width
'                Else
'                    objGroup.Left = obj.Left + bagSpac * (i - 1)
'                End If
'            ElseIf ((objsec.Left > obj.Left) And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2)) = -1) And (Abs(objsec.Left - obj.Left) < 300) And (Abs(objsec.Top - obj.Top) < 300) Or (objsec.Left < obj.Left And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2) = 1)) Then
'                'for bags going right-left
'                objGroup.Top = obj.Top
'                If Right(objGroup.ObjectName, 2) = "0" & bagNr Then
'                    objGroup.Left = obj.Left
'                ElseIf Right(objGroup.ObjectName, 2) = bagNr Then
'                    objGroup.Left = obj.Left
'                Else
'                    objGroup.Left = obj.Left + obj.Width - objGroup.Width - bagSpac * (i - 1)
'                End If
'            End If
'        ElseIf obj.RotationAngle = 270 Then
'            If ((objsec.Top > obj.Top) And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2)) = -1) And (Abs(objsec.Left - obj.Left) < 300) And (Abs(objsec.Top - obj.Top) < 300) Or (objsec.Top < obj.Top And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2) = 1)) Then
'                'for bags going down-up
'                objGroup.Left = obj.Left + (obj.Width - objGroup.Width) / 2
'                If Right(objGroup.ObjectName, 2) = "0" & bagNr Then
'                    objGroup.Top = obj.Top + obj.Height / 2 - obj.Width / 2
'                ElseIf Right(objGroup.ObjectName, 2) = bagNr Then
'                    objGroup.Top = obj.Top + obj.Height / 2 - obj.Width / 2
'                Else
'                    objGroup.Top = (obj.Top - obj.Height / 2 + obj.Width / 2) - bagSpac * (i - 1)
'                End If
'            ElseIf ((objsec.Top < obj.Top) And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2)) = -1) And (Abs(objsec.Left - obj.Left) < 300) And (Abs(objsec.Top - obj.Top) < 300) Or (objsec.Top > obj.Top And (Right(objsec.ObjectName, 2) - Right(obj.ObjectName, 2) = 1)) Then
'                'for bags going up-down
'                objGroup.Left = obj.Left + (obj.Width - objGroup.Width) / 2
'                If Right(objGroup.ObjectName, 2) = "0" & bagNr Then
'                    objGroup.Top = obj.Top - obj.Height / 2 + obj.Width / 2
'                ElseIf Right(objGroup.ObjectName, 2) = bagNr Then
'                    objGroup.Top = obj.Top - obj.Height / 2 + obj.Width / 2
'                Else
'                    objGroup.Top = (obj.Top + obj.Height / 2 - obj.Width / 2) + bagSpac * (i - 1)
'                End If
'            End If
'        End If
'        End If
'        End If
'        End If
'        Next

    'add the mouse action click script
    If InStr(UCase(objGroup.ObjectName), "BAG_MAP_POS") > 0 Then
        If objGroup.Layer = 3 Then
            Set objCScript = objGroup.Events(1).Actions.AddAction(hmiActionCreationTypeCScript)
            objCScript.sourceCode = "char *mainWindow = GetMainWindow(lpszPictureName); SetTagChar(""BAG_NAME"", lpszObjectName); Plc4BagPopup(); DisplayPopupNextToObject(mainWindow, lpszPictureName, ""BagPopup Picture Window"", ""Plc4BagPopup"", lpszObjectName, -500,0);"
        End If
    End If

    Next

ElseIf (obj.Layer = 2 And obj.Type = "HMIPieSegment") Then

    i = 0
    convName = obj.ObjectName
    If obj.Type = "HMIPieSegment" Then
        If obj.StartAngle = 0 And obj.StartAngle > obj.EndAngle Then
        bagNr = (2 * pi / (360 / Abs(360 - obj.EndAngle)) * obj.Radius) / bagDensity   '40 cm in pixels is approx 13 px
        ElseIf obj.EndAngle = 0 And obj.StartAngle > obj.EndAngle Then
        bagNr = (2 * pi / (360 / Abs(obj.StartAngle - 360)) * obj.Radius) / bagDensity   '40 cm in pixels is approx 13 px
        Else
        bagNr = (2 * pi / (360 / Abs(obj.StartAngle - obj.EndAngle)) * obj.Radius) / bagDensity   '40 cm in pixels is approx 13 px
        End If
    End If
    If bagNr < 3 Then
        bagNr = 3
    ElseIf bagNr > 20 Then
        bagNr = 20
    Else
        bagNr = bagNr
    End If


    For i = 1 To bagNr

    If i < 10 Then
        bagName = convName & "_BAG_MAP_POS_0" & i
    Else
        bagName = convName & "_BAG_MAP_POS_" & i
    End If

    section = "Any object"

    'create the circle for the bag
    strCode = "long backgroundColour = GetBagStatusColours(lpszObjectName);  return backgroundColour;"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
       'objCircle.ObjectName = obj.ObjectName & "_BAG_MAP_POS_"
       objCircle.Left = 0
       objCircle.Top = 0
       objCircle.Radius = 14
       objCircle.GlobalColorScheme = 0
       objCircle.backColor = backColorBag
       Set objDynDialog = objCircle.Visible.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'BAG_VIS' && '" & bagName & "' > 0")
            With objDynDialog
                .ResultType = hmiResultTypeBool
                .BinaryResultInfo.NegativeValue = 0
                .BinaryResultInfo.PositiveValue = 1
            End With
        Set objCScript = objCircle.backColor.CreateDynamic(hmiDynamicCreationTypeCScript)
            With objCScript
            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
                .sourceCode = "long backgroundColour = GetBagStatusColours(""" & bagName & """);  return backgroundColour;"
            End With

      'create the upper output field
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField", "HMIIOField")
        objIOField.Left = objCircle.Left
        objIOField.Top = objCircle.Top + 4
        objIOField.Width = 30
        objIOField.Height = 10
        objIOField.FONTNAME = Arial
        objIOField.BoxType = Output
        objIOField.OutputFormat = "099"
        objIOField.LimitMin = -1.79769313486231E+308
        objIOField.LimitMax = 1.79769313486231E+308
        objIOField.backColor = RGB(175, 171, 176)
        objIOField.FillStyle = 65536 'transparent
        objIOField.BorderWidth = 0 'borderweight = 0
        objIOField.FONTBOLD = 1 'font bold
        objIOField.FONTSIZE = 11 'font size 11
        objIOField.AlignmentLeft = 1
        objIOField.AlignmentTop = 1
        objIOField.GlobalColorScheme = 0 'global color  scheme no
        Set objCScript = objIOField.OutputValue.CreateDynamic(hmiDynamicCreationTypeCScript)
            With objCScript
            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
                .sourceCode = "return(((unsigned long) GetTagDouble(""" & bagName & """) % 100000) / 1000);"
            End With

      'create the upper output field
    Set objIOField2 = ActiveDocument.HMIObjects.AddHMIObject("IOField", "HMIIOField")
        objIOField2.Left = objCircle.Left
        objIOField2.Top = objCircle.Top + 13
        objIOField2.Width = 30
        objIOField2.Height = 10
        objIOField2.FONTNAME = Arial
        objIOField2.BoxType = Output
        objIOField2.OutputFormat = "0999"
        objIOField2.LimitMin = -1.79769313486231E+308
        objIOField2.LimitMax = 1.79769313486231E+308
        objIOField2.backColor = RGB(175, 171, 176)
        objIOField2.FillStyle = 65536 'transparent
        objIOField2.BorderWidth = 0 'borderweight = 0
        objIOField2.FONTBOLD = 1 'font bold
        objIOField2.FONTSIZE = 11 'font size 11
        objIOField2.AlignmentLeft = 1
        objIOField2.AlignmentTop = 1
        objIOField2.GlobalColorScheme = 0 'global color  scheme no
        Set objCScript = objIOField2.OutputValue.CreateDynamic(hmiDynamicCreationTypeCScript)
            With objCScript
            Set objVarTrigger = .Trigger.VariableTriggers.Add(bagName, hmiVariableCycleTypeUponChange)
                .sourceCode = "return((unsigned long) GetTagDouble(""" & bagName & """) % 1000);"
            End With

    'group the created items
    With objCircle
        .Selected = True
    End With
    With objIOField
        .Selected = True
    End With
    With objIOField2
        .Selected = True
    End With


    Set objGroup = ActiveDocument.Selection.CreateGroup
        objGroup.ObjectName = bagName
        objGroup.Layer = 3
        objGroup.Top = obj.Top
        objGroup.Left = obj.Left




        If obj.StartAngle = 0 And obj.StartAngle > obj.EndAngle Then
            bagAngle = Abs(360 - obj.EndAngle) / (bagNr - 1)
        ElseIf obj.EndAngle = 0 And obj.StartAngle > obj.EndAngle Then
            bagAngle = Abs(obj.StartAngle - 360) / (bagNr - 1)
        Else
            bagAngle = (Abs(obj.StartAngle - obj.EndAngle + 4) / (bagNr - 1)) ' * pi / 180 * (i - 1)
        End If

        For Each objsec In objs
        If (objsec.Layer = 1 Or objsec.Layer = 2) And ((objsec.Left - obj.Left < 300) And (objsec.Top - obj.Top < 300)) Then
        If (Left(objsec.ObjectName, 2) = Left(obj.ObjectName, 2)) Then
            If (obj.StartAngle = 90 And obj.EndAngle = 180) Then
                'bags go left
                If ((objsec.Left < obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1)) Or (objsec.Top < obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1))) Then
                    objGroup.Top = obj.Top + (obj.Height - objGroup.Height - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                    objGroup.Left = obj.Left + (obj.Radius - objGroup.Height - 4) * Cos(bagAngle * (i - 1) * pi / 180)
                'bags go right
                ElseIf ((objsec.Left < obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1)) Or (objsec.Top < obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1))) Then
                    objGroup.Top = obj.Top + (obj.Height - objGroup.Height - 4) * Cos(bagAngle * (i - 1) * pi / 180)
                    objGroup.Left = obj.Left + (obj.Radius - objGroup.Height - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                End If
            ElseIf (obj.StartAngle = 180 And obj.EndAngle = 270) Then
                'bags go left
                If ((objsec.Left > obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1)) Or (objsec.Top < obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1))) Then 'obj.ObjectName = "D409" Then
                    objGroup.Top = obj.Top + (obj.Radius - objGroup.Height - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                    objGroup.Left = obj.Left + obj.Radius - objGroup.Width - (objGroup.Width - 2) * Cos(bagAngle * (i - 1) * pi / 180)
                'bags go right
                ElseIf ((objsec.Left > obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1)) Or (objsec.Top < obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1))) Then 'obj.ObjectName = "HBS410" Then
                    objGroup.Top = obj.Top + obj.Radius - obj.Height + (objGroup.Height - 2) * Cos(bagAngle * (i - 1) * pi / 180)
                    objGroup.Left = obj.Left + obj.Radius + objGroup.Width - obj.Width - (objGroup.Width - 2) * Sin(bagAngle * (i - 1) * pi / 180)
                End If
            ElseIf (obj.StartAngle = 270 And obj.EndAngle = 0) Then
                'bags go left
                If ((objsec.Left > obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1)) Or (objsec.Top > obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1))) Then
                    objGroup.Left = obj.Left - objGroup.Width + (obj.Radius) - (objGroup.Width - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                    objGroup.Top = obj.Top - objGroup.Height + (obj.Radius) - (objGroup.Height - 4) * Cos(bagAngle * (i - 1) * pi / 180)
                'bags go right
                ElseIf ((objsec.Left > obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1)) Or (objsec.Top > obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1))) Then
                    objGroup.Left = obj.Left - objGroup.Width + (obj.Radius) - (objGroup.Width - 4) * Cos(bagAngle * (i - 1) * pi / 180)
                    objGroup.Top = obj.Top - objGroup.Height + (obj.Radius) - (objGroup.Height - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                End If
            ElseIf (obj.StartAngle = 0 And obj.EndAngle = 90) Then
                'bags go left
                If ((objsec.Left < obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1)) Or (objsec.Top > obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1))) Then
                    objGroup.Left = obj.Left + (obj.Radius - objGroup.Width - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                    objGroup.Top = obj.Top - objGroup.Height + obj.Height - (obj.Radius - objGroup.Width - 4) * Cos(bagAngle * (i - 1) * pi / 180)
                'bags go right
                ElseIf ((objsec.Left < obj.Left And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = 1)) Or (objsec.Top > obj.Top And (CInt(Right(objsec.ObjectName, 2)) - CInt(Right(obj.ObjectName, 2)) = -1))) Then
                    objGroup.Left = obj.Left + (obj.Radius - objGroup.Width - 4) * Cos(bagAngle * (i - 1) * pi / 180)
                    objGroup.Top = obj.Top - objGroup.Height + obj.Height - (obj.Radius - objGroup.Width - 4) * Sin(bagAngle * (i - 1) * pi / 180)
                End If
            End If
        End If
        End If
        Next

    'add the mouse action click script
    If InStr(UCase(objGroup.ObjectName), "BAG_MAP_POS") > 0 Then
        If objGroup.Layer = 3 Then
            Set objCScript = objGroup.Events(1).Actions.AddAction(hmiActionCreationTypeCScript)
            objCScript.sourceCode = "SetTagChar(""BAG_NAME"", lpszObjectName); BagPopup (); SetPictureName(""Title.pdl"",""Control Picture Window"", ""BagPopup_New.pdl"");"
        End If
    End If

    Next

'End If
End If
Next

  GoTo EndOfScript2

ErrorHandler:
  If Err.Number <> 0 Then
    Msg = "Error # " & Str(Err.Number) & " on line " & Erl & " section '" & section & "' was generated by " _
           & Err.Source & Chr(13) & Err.Description
    MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
    If Not (obj Is Nothing) Then
      MsgBox "ogj = " + obj.ObjectName + " at pos (left,to) =" + Str(obj.Left) + "," + Str(obj.Top)
    Else
      MsgBox "obj is nothing"
    End If
  End If
  Resume Next
EndOfScript2:

End Sub
