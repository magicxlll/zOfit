VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShapeBtn"
Option Explicit

Public WithEvents Img As MSForms.Image
Attribute Img.VB_VarHelpID = -1
Public ParentForm As Object
Public ShapeCode As String

Private Sub Img_Click()
    ParentForm.OnShapeClicked ShapeCode
End Sub

Private Sub Img_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ParentForm.OnShapeDblClicked ShapeCode
End Sub
