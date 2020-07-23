Option Explicit
Sub pageFlip()

Dim obj As HMIObject
Dim dc3 As HMIObject
Dim objGr As HMIGroup
Dim objGrpd As HMIGroup
Dim i As Integer
Dim corr As Integer
Dim grp As HMIGroup
Dim objs As HMIObjects
Dim grpCol As New Collection
Dim mainCol As New Collection

Set objs = ActiveDocument.HMIObjects

Set dc3 = objs("@DataClass3")

corr = 2

Set objs = ActiveDocument.HMIObjects

'getting cols
For Each obj In objs
  If obj.Type <> "HMIGroup" And obj.GroupParent Is Nothing Then mainCol.Add obj, obj.ObjectName
Next
For Each obj In objs
  If obj.Type = "HMIGroup" And obj.GroupParent Is Nothing Then grpCol.Add obj, obj.ObjectName
Next

'performing actions
For Each grp In grpCol
  'move group
  If grp.Layer <> 29 And grp.ObjectName <> "@DataClass3" Then
      grp.left = (dc3.left - grp.left - grp.Width) + corr
  End If
Next
For Each obj In mainCol
  'move simple objects not contained by groups or not groups themvelves
  If obj.Layer <> 29 And obj.ObjectName <> "@DataClass3" Then
      obj.left = (dc3.left - obj.left - obj.Width) + corr
  End If
Next

'Call flipArrows
'
'Call lineEnds

MsgBox "Finished flipping"

End Sub
Sub flipArrows()

Dim obj As HMIObject
Dim objs As HMIObjects

Set objs = ActiveDocument.HMIObjects

For Each obj In objs
    If obj.Height = 10 And obj.Width = 8 And obj.Type = "HMIPolygon" Then
        ActiveDocument.selection.DeselectAll
        obj.Selected = True
        ActiveDocument.selection.FlipVertically
    End If
Next


End Sub
Sub lineEnds()

Dim obj As HMIObject
Dim objs As HMIObjects

Set objs = ActiveDocument.HMIObjects

For Each obj In objs
    If obj.Type = "HMILine" Then
'        If obj.ObjectName = "Line28" Then
            If obj.BorderEndStyle = 131072 And obj.Height = 0 Then
                obj.BorderEndStyle = 2
            ElseIf obj.BorderEndStyle = 2 And obj.Height = 0 Then
                obj.BorderEndStyle = 131072
            End If
'        End If
    End If
Next

End Sub

