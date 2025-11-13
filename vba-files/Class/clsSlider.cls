VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class: clsSlider
' Author: VBATools
' Version: 1.0.1
' Creation Date: 13.11.2025 10:59
' Last Update Date: 13.11.2025 15:57
' License: Apache License
'
' Description:
'    This class implements a custom slider control for VBA UserForms. The slider allows
'    users to select a value within a specified range by dragging a button along a track.
'    It supports both horizontal and vertical orientations and provides customizable
'    appearance and behavior options.
'
' Features:
'    - Horizontal and vertical orientation support
'    - Configurable min/max values and initial value
'    - Customizable appearance (colors, icons, formatting)
'    - Visible/hidden value label with positioning options
'    - Event handling for value changes and clicks
'    - Lock functionality to prevent user interaction
'    - Automatic positioning of value label (top, left, right, bottom)
'
' Usage:
'    1. Create an instance of clsSlider
'    2. Call Initialize method with appropriate parameters
'    3. Handle Change and Click events as needed
'    4. Access Value property to get/set current value
'
' Dependencies:
'    - MSForms library for Label controls
'    - XlRgbColor constants for color properties
'
'========================================================================================

Public Enum PositionValue
    Top = 1
    Left
    Right
    Bottom
End Enum

Private WithEvents mLabelBtn As MSForms.Label
Attribute mLabelBtn.VB_VarHelpID = -1
Private mLabelLineEmpty As MSForms.Label
Private mLabelLineFull As MSForms.Label
Private mLabelValue As MSForms.Label

Public Event Click(ByRef control As Object, ByRef Value As Single)
Public Event Change(ByRef control As Object, ByRef Value As Single)

Private mNameLabel  As String
Private mFormatValue As String
Private mValue      As Single
Private mMaxValue   As Single
Private mMinValue   As Single
Private snMin       As Single
Private snMax       As Single
Private mIsHorizontal As Boolean
Private mLocked     As Boolean
Private mPositionValue As PositionValue

Public Property Get Version() As String
    Version = "Class: clsSlider" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.1" & vbNewLine & _
            "Date of creation: 13.11.2025 10:59" & vbNewLine & _
            "Date of update: 13.11.2025 15:57" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description:"
End Property

Public Property Get Locked() As Boolean
    Locked = mLocked
End Property

Public Property Let Locked(ByVal Value As Boolean)
    mLocked = Value
End Property

Public Property Get Enabled() As Boolean
    Enabled = mLabelBtn.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    mLabelBtn.Enabled = Value
    mLabelValue.Enabled = Value
End Property

Public Property Get Visible() As Boolean
    Visible = mLabelBtn.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    mLabelBtn.Visible = Value
    mLabelLineEmpty.Visible = Value
    mLabelLineFull.Visible = Value
    mLabelValue.Visible = Value
End Property

Public Property Get VisibleLabelValue() As Boolean
    VisibleLabelValue = mLabelValue.Visible
End Property

Public Property Let VisibleLabelValue(ByVal Value As Boolean)
    mLabelValue.Visible = Value
End Property

Public Property Get Icon() As Long
    Icon = VBA.CLng(mLabelBtn.Caption)
End Property

Public Property Let Icon(ByVal Value As Long)
    mLabelBtn.Caption = VBA.ChrW(Value)
End Property

Public Property Get PositionLabelValue() As PositionValue
    PositionLabelValue = mPositionValue
End Property

Public Property Let PositionLabelValue(ByVal Value As PositionValue)
    mPositionValue = Value
    Call changePositionLabelValue(mPositionValue)
End Property

Public Property Get FormatValue() As String
    FormatValue = mFormatValue
End Property

Public Property Let FormatValue(ByVal Value As String)
    mFormatValue = Value
    If mFormatValue <> vbNullString Then
        mLabelValue.Caption = VBA.Format$(getValue(), mFormatValue)
    Else
        mLabelValue.Caption = getValue()
    End If
End Property

Public Property Get ForeColorValue() As XlRgbColor
    ForeColorValue = mLabelValue.ForeColor
End Property

Public Property Let ForeColorValue(ByVal Value As XlRgbColor)
    mLabelValue.ForeColor = Value
End Property

Public Property Get ForeColorBtn() As XlRgbColor
    ForeColorBtn = mLabelBtn.ForeColor
End Property

Public Property Let ForeColorBtn(ByVal Value As XlRgbColor)
    mLabelBtn.ForeColor = Value
End Property

Public Property Get BackColorFull() As XlRgbColor
    BackColorFull = mLabelLineFull.BackColor
End Property

Public Property Let BackColorFull(ByVal Value As XlRgbColor)
    mLabelLineFull.BackColor = Value
End Property

Public Property Get BackColorEmpty() As XlRgbColor
    BackColorEmpty = mLabelLineEmpty.BackColor
End Property

Public Property Let BackColorEmpty(ByVal Value As XlRgbColor)
    mLabelLineEmpty.BackColor = Value
End Property

Public Property Get Name() As String
    Name = mNameLabel
End Property

Public Property Get IsHorizontal() As Boolean
    IsHorizontal = mIsHorizontal
End Property

Public Property Get MinValue() As Single
    MinValue = mMinValue
End Property

Public Property Let MinValue(ByVal Value As Single)
    mMinValue = Value
End Property

Public Property Get MaxValue() As Single
    MaxValue = mMaxValue
End Property

Public Property Let MaxValue(ByVal Value As Single)
    mMaxValue = Value
End Property

Public Property Get Value() As Single
    Value = getValue()
End Property

Public Property Let Value(ByVal Value As Single)
    Select Case Value
        Case mMinValue To mMaxValue:
            Call SetValue(Value)
    End Select
End Property

Public Sub Initialize(ByRef labelSlider As MSForms.Label, _
        ByVal Value As Single, _
        ByVal MinValue As Single, _
        ByVal MaxValue As Single, _
        ByVal VisibleLabelValue As Boolean, _
        Optional FormatValue As String = vbNullString, _
        Optional PositionValue As PositionValue = Top, _
        Optional BackColorEmpty As XlRgbColor = rgbLightGray, _
        Optional BackColorFull As XlRgbColor = rgbGreenYellow, _
        Optional ForeColorBtn As XlRgbColor = rgbBlack, _
        Optional ForeColorValue As XlRgbColor = rgbBlack, _
        Optional Icon As Long = 59963)

    Set mLabelLineEmpty = labelSlider
    mValue = Value
    mMaxValue = MaxValue
    mMinValue = MinValue
    mFormatValue = FormatValue
    mPositionValue = PositionValue

    With mLabelLineEmpty
        mIsHorizontal = .Width / .Height > 1
        If mIsHorizontal Then
            .Height = 3
            snMin = .Left
            snMax = snMin + .Width
        Else
            .Width = 3
            snMin = .Top
            snMax = snMin + .Height
        End If
        .Caption = vbNullString
        mNameLabel = .Name
        .BackColor = BackColorEmpty
        Set mLabelLineFull = mLabelLineEmpty.Parent.Controls.Add("Forms.label.1", mNameLabel & "_full", True)
        Set mLabelBtn = mLabelLineEmpty.Parent.Controls.Add("Forms.label.1", mNameLabel & "_btn", True)
        Set mLabelValue = mLabelLineEmpty.Parent.Controls.Add("Forms.label.1", mNameLabel & "_value", True)
    End With

    With mLabelLineFull
        .Caption = vbNullString
        .Top = mLabelLineEmpty.Top
        .Left = mLabelLineEmpty.Left
        .BackColor = BackColorFull
        .ZOrder 0
    End With

    With mLabelBtn
        .Height = 12
        .Width = .Height
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe MDL2 Assets"
        .Font.Size = 10
        .Caption = VBA.ChrW(Icon)
        .ForeColor = ForeColorBtn
        .ZOrder 0
    End With

    With mLabelValue
        .Visible = VisibleLabelValue
        .BackStyle = fmBackStyleTransparent
        .Font.Name = mLabelLineEmpty.Font.Name
        .Font.Size = mLabelLineEmpty.Font.Size
        .TextAlign = fmTextAlignCenter
        .ForeColor = ForeColorValue
        .ZOrder 0
    End With
    Call SetValue(mValue)
End Sub

Private Sub mLabelBtn_Click()
    If mLocked Then Exit Sub
    RaiseEvent Click(mLabelBtn, mValue)
End Sub

Private Sub mLabelBtn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button <> 1 Or mLocked Then Exit Sub
    With mLabelBtn
        If mIsHorizontal Then
            .Left = mLabelLineFull.Left + mLabelLineFull.Width - .Width / 2
        Else
            .Top = mLabelLineFull.Top + mLabelLineFull.Height - .Height / 2
        End If
    End With
End Sub

Private Sub mLabelBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button <> 1 Or mLocked Then Exit Sub
    With mLabelBtn
        If mIsHorizontal Then
            Select Case .Left + x
                Case snMin - 0.5 To snMax + 0.5:
                    .Left = .Left + x
                    Select Case mLabelBtn.Left + x
                        Case Is < snMin
                            mLabelLineFull.Width = 0
                        Case Is > snMax
                            mLabelLineFull.Width = mLabelLineEmpty.Width
                        Case Else
                            If mLabelLineFull.Width + x < 0 Then
                                mLabelLineFull.Width = 0
                            Else
                                mLabelLineFull.Width = mLabelLineFull.Width + x
                            End If
                    End Select
            End Select
        Else
            Select Case .Top + y
                Case snMin - 0.5 To snMax + 0.5:
                    .Top = .Top + y
                    Select Case mLabelBtn.Top + y
                        Case Is < snMin
                            mLabelLineFull.Height = 0
                        Case Is > snMax
                            mLabelLineFull.Height = mLabelLineEmpty.Height
                        Case Else
                            If mLabelLineFull.Height + y < 0 Then
                                mLabelLineFull.Height = 0
                            Else
                                mLabelLineFull.Height = mLabelLineFull.Height + y
                            End If
                    End Select
            End Select
        End If
    End With
    If mFormatValue = vbNullString Then
        mLabelValue.Caption = getValue()
    Else
        mLabelValue.Caption = VBA.Format$(getValue(), mFormatValue)
    End If
    RaiseEvent Change(mLabelBtn, mValue)
End Sub

Private Function getValue() As Single
    If mIsHorizontal Then
        getValue = mMinValue + mLabelLineFull.Width / mLabelLineEmpty.Width * (mMaxValue - mMinValue)
    Else
        getValue = mMinValue + mLabelLineFull.Height / mLabelLineEmpty.Height * (mMaxValue - mMinValue)
    End If
    If getValue > mMaxValue Then
        getValue = mMaxValue
    ElseIf getValue < mMinValue Then
        getValue = mMinValue
    End If
    mValue = getValue
End Function

Private Sub SetValue(ByRef snValue As Single)
    mValue = snValue
    With mLabelLineFull
        If mIsHorizontal Then
            .Height = mLabelLineEmpty.Height
            .Width = (mValue - mMinValue) / (mMaxValue - mMinValue) * mLabelLineEmpty.Width
        Else
            .Height = (mValue - mMinValue) / (mMaxValue - mMinValue) * mLabelLineEmpty.Height
            .Width = mLabelLineEmpty.Width
        End If
    End With
    With mLabelBtn
        If mIsHorizontal Then
            .Top = mLabelLineEmpty.Top + mLabelLineEmpty.Height / 2 - .Height / 2 + 0.5
            .Left = mLabelLineFull.Left + mLabelLineFull.Width - .Width / 2
        Else
            .Top = mLabelLineFull.Top + mLabelLineFull.Height - .Height / 2
            .Left = mLabelLineEmpty.Left + mLabelLineEmpty.Width / 2 - .Width / 2
        End If
    End With
    With mLabelValue
        .Height = 10
        .Width = 30
        Call changePositionLabelValue(mPositionValue)
        .Caption = mValue
        If mFormatValue <> vbNullString Then .Caption = VBA.Format$(mValue, mFormatValue)
    End With
End Sub

Private Sub changePositionLabelValue(ByVal pos As PositionValue)
    With mLabelValue
        Select Case pos
            Case PositionValue.Top:
                .Top = mLabelLineEmpty.Top - .Height - 5
                .Left = mLabelLineEmpty.Left + mLabelLineEmpty.Width / 2 - .Width / 2
            Case PositionValue.Left:
                .Top = mLabelLineEmpty.Top + mLabelLineEmpty.Height / 2 - .Height / 2
                .Left = mLabelLineEmpty.Left - .Width
            Case PositionValue.Right:
                .Top = mLabelLineEmpty.Top + mLabelLineEmpty.Height / 2 - .Height / 2
                .Left = mLabelLineEmpty.Left + mLabelLineEmpty.Width + 5
            Case PositionValue.Bottom:
                .Top = mLabelLineEmpty.Top + mLabelLineEmpty.Height + 5
                .Left = mLabelLineEmpty.Left + mLabelLineEmpty.Width / 2 - .Width / 2
        End Select
    End With
End Sub

Private Sub Class_Terminate()
    Set mLabelBtn = Nothing
    Set mLabelLineEmpty = Nothing
    Set mLabelLineFull = Nothing
    Set mLabelValue = Nothing
End Sub