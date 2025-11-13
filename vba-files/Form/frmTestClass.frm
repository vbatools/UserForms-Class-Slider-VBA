VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestClass 
   Caption         =   "Test class:"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "frmTestClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private WithEvents clsSL As clsSlider
Attribute clsSL.VB_VarHelpID = -1
Private WithEvents clsSL2 As clsSlider
Attribute clsSL2.VB_VarHelpID = -1

Private Sub btnColor_Click()
    clsSL.ForeColorBtn = rgbOrange
End Sub

Private Sub btnEmptyColor_Click()
    clsSL.BackColorEmpty = rgbGrey
End Sub

Private Sub btnFullColor_Click()
    clsSL.BackColorFull = rgbIndianRed
End Sub

Private Sub btnValueColor_Click()
    clsSL.ForeColorValue = rgbOrchid
End Sub

Private Sub chbFormat_Click()
    With clsSL2
        If .FormatValue = vbNullString Then
            .FormatValue = "##%"
        Else
            .FormatValue = vbNullString
        End If
    End With
End Sub

Private Sub chbLocked_Click()
   clsSL.Locked = Not clsSL.Locked
End Sub

Private Sub chbVisible_Click()
    clsSL.Visible = Not clsSL.Visible
End Sub

Private Sub chbEnabled_Click()
    clsSL.Enabled = Not clsSL.Enabled
End Sub

Private Sub clsSL_Click(control As Object, Value As Single)
    lbClick.Caption = "Click control: " & control.Name & vbNewLine & "Value: " & Value & vbNewLine & "Max: " & clsSL.MaxValue & vbNewLine & "Min: " & clsSL.MinValue
End Sub

Private Sub clsSL2_Change(control As Object, Value As Single)
    lbClick.Caption = "Click control: " & control.Name & vbNewLine & "Value: " & Value & vbNewLine & "Max: " & clsSL2.MaxValue & vbNewLine & "Min: " & clsSL2.MinValue
End Sub

Private Sub optBotton_Click()
    clsSL2.PositionLabelValue = PositionValue.Bottom
    clsSL.PositionLabelValue = PositionValue.Bottom
End Sub

Private Sub optLeft_Click()
    clsSL2.PositionLabelValue = PositionValue.Left
    clsSL.PositionLabelValue = PositionValue.Left
End Sub

Private Sub optRight_Click()
    clsSL2.PositionLabelValue = PositionValue.Right
    clsSL.PositionLabelValue = PositionValue.Right
End Sub

Private Sub optTOP_Click()
    clsSL2.PositionLabelValue = PositionValue.Top
    clsSL.PositionLabelValue = PositionValue.Top
End Sub

Private Sub txtSetValue_Change()
    clsSL.Value = Val(txtSetValue.Value)
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With

    Set clsSL = New clsSlider
    Set clsSL2 = New clsSlider
    With clsSL
        Call .Initialize(Label1, 60, 50, 100, True, , PositionValue.Bottom)
    End With
    With clsSL2
        Call .Initialize(Label2, 0.5, 0, 1, True, "##%", PositionValue.Right)
    End With

    lbVersion.Caption = clsSL.Version
End Sub