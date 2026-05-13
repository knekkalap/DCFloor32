Attribute VB_Name = "Module1"
Option Explicit

      Sub Resize_For_Resolution(ByVal SFX As Single, _
       ByVal SFY As Single, MyForm As Form)
      Dim I As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' average scale
      ' Size the Controls for the new resolution
      On Error Resume Next  ' for read-only or nonexistent properties
      With MyForm
        For I = 0 To .Count - 1
         If TypeOf .Controls(I) Is ComboBox Then   ' cannot change Height
           .Controls(I).Left = .Controls(I).Left * SFX
           .Controls(I).Top = .Controls(I).Top * SFY
           .Controls(I).Width = .Controls(I).Width * SFX
         Else
           .Controls(I).Move .Controls(I).Left * SFX, _
            .Controls(I).Top * SFY, _
            .Controls(I).Width * SFX, _
            .Controls(I).Height * SFY
         End If
           ' Be sure to resize and reposition before changing the FontSize
           .Controls(I).FontSize = .Controls(I).FontSize * SFFont
        Next I
        If RePosForm Then
          ' Now size the Form
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
      End Sub


