<div align="center">

## SelectListBoxes


</div>

### Description

Move selections back and forth between 2 listboxes. You can use command buttons or double click selections. This sample uses printer font names for testing.
 
### More Info
 
Form name frmSelectList should contain 2 listboxes named lstLists in a control array and 4 command buttons named cmdArrows in a control array that are used for arrows.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert Sadler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-sadler.md)
**Level**          |Beginner
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-sadler-selectlistboxes__1-3667/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub cmdArrows_Click(Index As Integer)
 Dim I As Integer
  Select Case Index
   Case 0     ' > Button
     For i = 0 To lstLists(0).ListCount - 1
      If lstLists(0).Selected(i) Then
        lstLists(1).AddItem lstLists(0).List(i)
        lstLists(1).ItemData(lstLists(1).NewIndex) = lstLists(0).ItemData(i)
      End If
     Next i
     For i = (lstLists(0).ListCount - 1) To 0 Step -1
      If lstLists(0).Selected(i) Then
        lstLists(0).RemoveItem i
      End If
     Next i
   Case 1     ' >> Button
     For i = 0 To lstLists(0).ListCount - 1
       lstLists(1).AddItem lstLists(0).List(i)
       lstLists(1).ItemData(lstLists(1).NewIndex) = lstLists(0).ItemData(i)
     Next i
     For i = (lstLists(0).ListCount - 1) To 0 Step -1
       lstLists(0).RemoveItem i
     Next i
   Case 2     ' < Button
     For i = 0 To lstLists(1).ListCount - 1
      If lstLists(1).Selected(i) Then
       lstLists(0).AddItem lstLists(1).List(i)
       lstLists(0).ItemData(lstLists(0).NewIndex) = lstLists(1).ItemData(i)
      End If
     Next i
     For i = (lstLists(1).ListCount - 1) To 0 Step -1
      If lstLists(1).Selected(i) Then
        lstLists(1).RemoveItem i
      End If
     Next i
   Case 3     ' << Button
     For i = 0 To lstLists(1).ListCount - 1
      lstLists(0).AddItem lstLists(1).List(i)
      lstLists(0).ItemData(lstLists(0).NewIndex) = lstLists(1).ItemData(i)
     Next i
     For i = (lstLists(1).ListCount - 1) To 0 Step -1
      lstLists(1).RemoveItem i
     Next i
 End Select
 SetButtons
End Sub
Private Sub Form_Load()
 Dim I As Integer, Flag As Boolean
 cmdArrows(0).Caption = ">"
 cmdArrows(1).Caption = ">>"
 cmdArrows(2).Caption = "<"
 cmdArrows(3).Caption = "<<"
 For I = 0 To Printer.FontCount - 1
 frmSelectList.lstLists(0).AddItem Printer.Fonts(I)
 Next I
 SetButtons ' go to set Select buttons
End Sub
Private Sub lstLists_Click(Index As Integer)
 SetButtons ' go to set select buttons
End Sub
Public Sub SetButtons()
 cmdArrows(0).Enabled = False
 cmdArrows(1).Enabled = False
 cmdArrows(2).Enabled = False
 cmdArrows(3).Enabled = False
 If lstLists(0).ListCount > 0 Then
 cmdArrows(1).Enabled = True ' >> Button
 If lstLists(0).SelCount > 0 Then
 cmdArrows(0).Enabled = True ' > Button
 End If
 End If
 If lstLists(1).ListCount > 0 Then
 cmdArrows(3).Enabled = True ' << Button
 If lstLists(1).SelCount > 0 Then
 cmdArrows(2).Enabled = True ' < Button
 End If
 End If
End Sub
Private Sub lstLists_DblClick(Index As Integer)
 Select Case Index
 Case 0
 cmdArrows_Click (0) ' > Button
 Case 1
 cmdArrows_Click (2) ' < Button
 End Select
End Sub
```

