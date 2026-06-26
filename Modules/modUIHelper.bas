Option Explicit

Public Sub CenterFormOnExcel(ByVal frm As Object)

    frm.StartUpPosition = 0

    frm.Top = Application.Top + (Application.Height - frm.Height) / 2
    frm.Left = Application.Left + (Application.Width - frm.Width) / 2

End Sub
