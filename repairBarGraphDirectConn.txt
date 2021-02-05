Sub repairDirectConnectionStatusObjects()

Dim obj As HMIObject
Dim ev As HMIEvent
Dim evs As HMIEvents
Dim str As String

Dim prop As HMIProperty
Dim props As HMIProperties

Dim what As HMIGraphicObject

Dim dyn As HMIDynamicDialog
Dim scr As HMIScriptInfo
Dim dir As HMIDirectConnection

Set objs = ActiveDocument.HMIObjects

For Each obj In objs

  If InStr(1, obj.ObjectName, "Status", vbTextCompare) > 0 Then
  
    Set props = obj.Properties
    
    For Each prop In props
    
      Rem **** to repair reference ****
Rem      If prop.Name = "Transparency" Then
Rem     If prop.Events.Item(1).Actions.Count > 0 Then
Rem       If prop.Events.Item(1).Actions(1).ActionType = hmiActionTypeDirectConnection Then
Rem         Set dir = prop.Events.Item(1).Actions(1)
Rem           With dir
Rem           .SourceLink.ObjectName = obj.ObjectName
Rem           .SourceLink.AutomationName = prop.Name
Rem           .DestinationLink.AutomationName = prop.Name
Rem           .DestinationLink.ObjectName = obj.ObjectName
Rem           End With
Rem       End If
Rem     End If
Rem     End If
      Rem **** to repair reference ****

      Rem ****** to delete ************
      If prop.Name = "Transparency" Then
      If prop.Events.Item(1).Actions.Count > 0 Then
        If prop.Events.Item(1).Actions(1).ActionType = hmiActionTypeDirectConnection Then
          On Error Resume Next
          prop.Events.Item(1).Actions.Item(1).Delete
          
        End If
      End If
      End If
      Rem ****** to delete ************
    Next
    
  End If

Next

End Sub
