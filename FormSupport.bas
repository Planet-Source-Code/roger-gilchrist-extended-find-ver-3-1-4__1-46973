Attribute VB_Name = "FormSupport"
'General services to support sub forms in Extended Find
Option Explicit

Public Sub Frame2Tab(tb As TabStrip, _
                     frms As Variant, _
                     CurFrame As Long)

  'assumes that the frames for the tabstrip are indexed consecitively from 1
  ' No need to change frame.

 If tb.SelectedItem.Index = CurFrame Then
  Exit Sub
 End If
 ' Otherwise, show new frame, hide old.
 With frms(tb.SelectedItem.Index)
  .Visible = True
  .Move tb.ClientLeft, tb.ClientTop
 End With
 If CurFrame > 0 Then
  If CurFrame <> tb.SelectedItem.Index Then
   frms(CurFrame).Visible = False
  End If
 End If
 ' Set new value.
 CurFrame = tb.SelectedItem.Index

End Sub

Public Sub LoadFormPosition(frm As Form)

  'Requires AppDetails to supply top of Registry branch
  'You could also hard code it if you want

 With frm
  .Left = GetSetting(AppDetails, "Settings", .Name & "Left", .Left)
  .Top = GetSetting(AppDetails, "Settings", .Name & "Top", .Top)
  If frm.BorderStyle = vbSizableToolWindow Or .BorderStyle = vbSizable Then
   'don't bother to load if form is not resizable
   .Width = GetSetting(AppDetails, "Settings", .Name & "Width", .Width)
   .Top = GetSetting(AppDetails, "Settings", .Name & "Height", .Height)
  End If
 End With 'Me

End Sub

Public Sub SaveFormPosition(frm As Form)

  'Requires AppDetails to supply top of Registry branch
  'You could also hard code it if you want

 With frm
  SaveSetting AppDetails, "Settings", .Name & "Left", .Left
  SaveSetting AppDetails, "Settings", .Name & "Top", .Top
  If .BorderStyle = vbSizableToolWindow Or .BorderStyle = vbSizable Then
   'don't bother to save if form is not resizable
   SaveSetting AppDetails, "Settings", .Name & "Width", .Width
   SaveSetting AppDetails, "Settings", .Name & "Height", .Height
  End If
 End With 'frm

End Sub

':) Roja's VB Code Fixer V1.1.4 (30/07/2003 1:25:28 PM) 2 + 60 = 62 Lines Thanks Ulli for inspiration and lots of code.

