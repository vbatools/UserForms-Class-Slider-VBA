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
' Version: 1.0.3
' Creation Date: 13.11.2025 10:59
' Last Update Date: 29.11.2025 14:00
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
    Top = 1  ' Position of the value label above the slider
    Left   ' Position of the value label to the left of the slider
    Right    ' Position of the value label to the right of the slider
    Bottom    ' Position of the value label below the slider
End Enum

Private WithEvents mTxtBtn As MSForms.TextBox
Attribute mTxtBtn.VB_VarHelpID = -1
Private WithEvents mParent As MSForms.UserForm
Attribute mParent.VB_VarHelpID = -1
Private mLabelLineEmpty As MSForms.Label
Private mLabelLineFull As MSForms.Label
Private mLabelValue As MSForms.Label

Public Event DblClick(ByRef control As Object, ByRef Value As Single)
Public Event Change(ByRef control As Object, ByRef Value As Single)

Private mNameLabel As String        ' Name of the main slider label
Private mFormatValue As String      ' Format for displaying the value
Private mValue As Single            ' Current slider value
Private mMaxValue As Single         ' Maximum slider value
Private mMinValue As Single         ' Minimum slider value
Private mSliderMinPosition As Single    ' Minimum slider position
Private mSliderMaxPosition As Single    ' Maximum slider position
Private mIsHorizontal As Boolean    ' Flag for horizontal orientation
Private mLocked As Boolean          ' Flag for locking the slider
Private mPositionValue As PositionValue  ' Position of the value label
Private mButtonColor As XlRgbColor
Private mStep As Single

Public Property Get Version() As String
    ' Returns version information about the class
    ' Example: Debug.Print slider.Version
    Version = "Class: clsSlider" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "Version: 1.0.3" & vbNewLine & _
            "Date of creation: 13.11.2025 10:59" & vbNewLine & _
            "Date of update: 29.11.2025 14:00" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Description: Custom slider control for VBA UserForms with horizontal/vertical orientation support"
End Property

Public Property Get StepSlider() As Single
    StepSlider = mStep
End Property

Public Property Let StepSlider(ByVal Value As Single)

    If Value = 0 Or (mMaxValue - mMinValue) < Value Then
        mStep = (mMaxValue - mMinValue) / mLabelLineEmpty.Width
    Else
        mStep = Value
    End If
End Property

Public Property Get Locked() As Boolean
    ' Returns the lock state of the slider
    ' If True, the slider does not respond to user actions
    ' Example: If slider.Locked Then Debug.Print "Slider is locked"
    Locked = mLocked
End Property

Public Property Let Locked(ByVal Value As Boolean)
    ' Sets the lock state of the slider
    ' When True, the slider does not respond to user actions
    ' Example: slider.Locked = True
    mLocked = Value
End Property

Public Property Get Enabled() As Boolean
    ' Returns the enabled state of the slider
    ' If False, the slider appears grayed out and does not respond to user actions
    ' Example: If slider.Enabled Then Debug.Print "Slider is enabled"
    Enabled = mTxtBtn.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    ' Sets the enabled state of the slider
    ' When False, the slider appears grayed out and does not respond to user actions
    ' Example: slider.Enabled = False
    mTxtBtn.Enabled = Value
    mLabelValue.Enabled = Value
End Property

Public Property Get Visible() As Boolean
    ' Returns the visibility state of all slider elements
    ' If False, all slider elements are hidden
    ' Example: If slider.Visible Then Debug.Print "Slider is visible"
    Visible = mTxtBtn.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    ' Sets the visibility state of all slider elements
    ' When False, all slider elements are hidden
    ' Example: slider.Visible = False
    mTxtBtn.Visible = Value
    mLabelLineEmpty.Visible = Value
    mLabelLineFull.Visible = Value
    mLabelValue.Visible = Value
End Property

Public Property Get VisibleLabelValue() As Boolean
    ' Returns the visibility state of the slider value label
    ' If False, the value label is not displayed
    ' Example: If slider.VisibleLabelValue Then Debug.Print "Value label is visible"
    VisibleLabelValue = mLabelValue.Visible
End Property

Public Property Let VisibleLabelValue(ByVal Value As Boolean)
    ' Sets the visibility state of the slider value label
    ' When False, the value label is hidden
    ' Example: slider.VisibleLabelValue = False
    mLabelValue.Visible = Value
End Property

Public Property Get Icon() As Long
    ' Returns the icon on the slider button (as a Unicode character value)
    ' Example: Dim currentIcon As Long: currentIcon = slider.Icon
    Icon = VBA.CLng(mTxtBtn.Caption)
End Property

Public Property Let Icon(ByVal Value As Long)
    ' Sets the icon on the slider button
    ' Accepts a Unicode character value
    ' Example: slider.Icon = 59963  (sets an icon from Segoe MDL2 Assets font)
    mTxtBtn.Value = VBA.ChrW(Value)
End Property

Public Property Get PositionLabelValue() As PositionValue
    ' Returns the current position of the slider value label
    ' Possible values: Top, Left, Right, Bottom
    ' Example: Dim pos As PositionValue: pos = slider.PositionLabelValue
    PositionLabelValue = mPositionValue
End Property

Public Property Let PositionLabelValue(ByVal Value As PositionValue)
    ' Sets the position of the slider value label
    ' Accepts a value from the PositionValue enumeration (Top, Left, Right, Bottom)
    ' Example: slider.PositionLabelValue = clsSlider.PositionValue.Right
    mPositionValue = Value
    Call UpdateValueLabelPosition(mPositionValue)
End Property

Public Property Get FormatValue() As String
    ' Returns the format for displaying the slider value
    ' Used to format the numeric value when displayed
    ' Example: Dim fmt As String: fmt = slider.FormatValue
    FormatValue = mFormatValue
End Property

Public Property Let FormatValue(ByVal Value As String)
    ' Sets the format for displaying the slider value
    ' Accepts a format string similar to that used in the Format function
    ' Example: slider.FormatValue = "0.00%"  (displays value as percentage)
    mFormatValue = Value
    If mFormatValue <> vbNullString Then
        mLabelValue.Caption = VBA.Format$(getValue(), mFormatValue)
    Else
        mLabelValue.Caption = getValue()
    End If
End Property

Public Property Get ForeColorValue() As XlRgbColor
    ' Returns the text color of the slider value label
    ' Example: Dim color As XlRgbColor: color = slider.ForeColorValue
    ForeColorValue = mLabelValue.ForeColor
End Property

Public Property Let ForeColorValue(ByVal Value As XlRgbColor)
    ' Sets the text color of the slider value label
    ' Example: slider.ForeColorValue = rgbRed
    mLabelValue.ForeColor = Value
End Property

Public Property Get ForeColorBtn() As XlRgbColor
    ' Returns the text color of the slider button
    ' Example: Dim color As XlRgbColor: color = slider.ForeColorBtn
    ForeColorBtn = mTxtBtn.ForeColor
End Property

Public Property Let ForeColorBtn(ByVal Value As XlRgbColor)
    ' Sets the text color of the slider button
    ' Example: slider.ForeColorBtn = rgbBlue
    mTxtBtn.ForeColor = Value
    mButtonColor = Value
End Property

Public Property Get BackColorFull() As XlRgbColor
    ' Returns the color of the filled part of the slider line
    ' Example: Dim color As XlRgbColor: color = slider.BackColorFull
    BackColorFull = mLabelLineFull.BackColor
End Property

Public Property Let BackColorFull(ByVal Value As XlRgbColor)
    ' Sets the color of the filled part of the slider line
    ' Example: slider.BackColorFull = rgbGreen
    mLabelLineFull.BackColor = Value
End Property

Public Property Get BackColorEmpty() As XlRgbColor
    ' Returns the color of the empty part of the slider line
    ' Example: Dim color As XlRgbColor: color = slider.BackColorEmpty
    BackColorEmpty = mLabelLineEmpty.BackColor
End Property

Public Property Let BackColorEmpty(ByVal Value As XlRgbColor)
    ' Sets the color of the empty part of the slider line
    ' Example: slider.BackColorEmpty = rgbGray
    mLabelLineEmpty.BackColor = Value
End Property

Public Property Get Name() As String
    ' Returns the name of the main slider label
    ' Example: Dim sliderName As String: sliderName = slider.Name
    Name = mNameLabel
End Property

Public Property Get IsHorizontal() As Boolean
    ' Returns whether the slider is horizontal
    ' True - horizontal, False - vertical
    ' Example: If slider.IsHorizontal Then Debug.Print "Slider is horizontal"
    IsHorizontal = mIsHorizontal
End Property

Public Property Get MinValue() As Single
    ' Returns the minimum value of the slider
    ' Example: Dim minVal As Single: minVal = slider.MinValue
    MinValue = mMinValue
End Property

Public Property Let MinValue(ByVal Value As Single)
    ' Sets the minimum value of the slider
    ' Example: slider.MinValue = 0
    mMinValue = Value
End Property

Public Property Get MaxValue() As Single
    ' Returns the maximum value of the slider
    ' Example: Dim maxVal As Single: maxVal = slider.MaxValue
    MaxValue = mMaxValue
End Property

Public Property Let MaxValue(ByVal Value As Single)
    ' Sets the maximum value of the slider
    ' Example: slider.MaxValue = 100
    mMaxValue = Value
End Property

Public Property Get Value() As Single
    ' Returns the current value of the slider
    ' Example: Dim currentVal As Single: currentVal = slider.Value
    Value = mValue
End Property

Public Property Let Value(ByVal Value As Single)
    ' Sets the current value of the slider
    ' The value must be within the range from MinValue to MaxValue
    ' Example: slider.Value = 50
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
        Optional StepSlider As Single = 0, _
        Optional FormatValue As String = vbNullString, _
        Optional PositionValue As PositionValue = Top, _
        Optional BackColorEmpty As XlRgbColor = rgbLightGray, _
        Optional BackColorFull As XlRgbColor = rgbGreenYellow, _
        Optional ForeColorBtn As XlRgbColor = rgbBlack, _
        Optional ForeColorValue As XlRgbColor = rgbBlack, _
        Optional Icon As Long = 59963)
    ' Initializes the slider with specified parameters
    ' Parameters:
    ' - labelSlider: the main label that will be used as the empty slider line
    ' - Value: initial value of the slider
    ' - MinValue: minimum value of the slider
    ' - MaxValue: maximum value of the slider
    ' - VisibleLabelValue: visibility of the value label
    ' - FormatValue: format for displaying the value (default vbNullString)
    ' - PositionValue: position of the value label (default Top)
    ' - BackColorEmpty: color of the empty part of the line (default rgbLightGray)
    ' - BackColorFull: color of the filled part of the line (default rgbGreenYellow)
    ' - ForeColorBtn: color of the button (default rgbBlack)
    ' - ForeColorValue: color of the value (default rgbBlack)
    ' - Icon: icon on the button (default 59963)
    '
    ' Usage example:
    ' Dim slider As New clsSlider
    ' Call slider.Initialize(Me.Label1, 50, 0, 100, True, "0.0", clsSlider.PositionValue.Top)
    '
    ' Check if the value range is valid
    If MinValue >= MaxValue Then
        Err.Raise vbObjectError + 1001, "clsSlider", "Minimum value must be less than maximum value"
    End If

    ' Check if the initial value is within the range
    If Value < MinValue Or Value > MaxValue Then
        Err.Raise vbObjectError + 1002, "clsSlider", "Initial value must be within the MinValue..MaxValue range"
    End If

    Set mLabelLineEmpty = labelSlider
    Set mParent = mLabelLineEmpty.Parent
    mValue = Value
    mMaxValue = MaxValue
    mMinValue = MinValue
    mFormatValue = FormatValue
    mPositionValue = PositionValue
    mButtonColor = ForeColorBtn

    If StepSlider = 0 Or (mMaxValue - mMinValue) < StepSlider Then
            mStep = (mMaxValue - mMinValue) / mLabelLineEmpty.Width
        Else
            mStep = StepSlider
        End If

    With mLabelLineEmpty
        mIsHorizontal = .Width / .Height > 1
        If mIsHorizontal Then
            .Height = 3
            mSliderMinPosition = .Left
            mSliderMaxPosition = mSliderMinPosition + .Width
        Else
            .Width = 3
            mSliderMinPosition = .Top
            mSliderMaxPosition = mSliderMinPosition + .Height
        End If
        .Caption = vbNullString
        mNameLabel = .Name
        .BackColor = BackColorEmpty
        Set mLabelLineFull = mLabelLineEmpty.Parent.Controls.Add("Forms.label.1", mNameLabel & "_full", True)
        Set mTxtBtn = mLabelLineEmpty.Parent.Controls.Add("Forms.TextBox.1", mNameLabel & "_btn", True)
        Set mLabelValue = mLabelLineEmpty.Parent.Controls.Add("Forms.label.1", mNameLabel & "_value", True)
    End With

    With mLabelLineFull
        .Caption = vbNullString
        .Top = mLabelLineEmpty.Top
        .Left = mLabelLineEmpty.Left
        .BackColor = BackColorFull
        .ZOrder 0
    End With

    With mTxtBtn
        .Height = 20
        .Width = .Height
        .BackStyle = fmBackStyleTransparent
        .TextAlign = fmTextAlignCenter
        .Font.Name = "Segoe MDL2 Assets"
        .Font.Size = 10
        .Value = VBA.ChrW(Icon)
        .Locked = True
        .SpecialEffect = fmSpecialEffectFlat
        '.BorderStyle = fmBorderStyleSingle
        .ForeColor = ForeColorBtn
        .MousePointer = fmMousePointerArrow
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

Private Sub mTxtBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Dim sShift      As Double
    If mIsHorizontal Then
        Select Case KeyCode
            Case 37, 65
                Call ApplyHoverEffect(100)
                sShift = mValue - mStep
                If sShift < mMinValue Then sShift = mMinValue
                Me.Value = sShift
            Case 39, 68
                Call ApplyHoverEffect(100)
                sShift = mValue + mStep
                If sShift > mMaxValue Then sShift = mMaxValue
                Me.Value = sShift
        End Select
    Else
        Select Case KeyCode
            Case 38, 87
                Call ApplyHoverEffect(100)
                KeyCode = 87
                sShift = mValue - mStep
                If sShift < mMinValue Then sShift = mMinValue
                Me.Value = sShift
            Case 40, 83
                Call ApplyHoverEffect(100)
                KeyCode = 83
                sShift = mValue + mStep
                If sShift > mMaxValue Then sShift = mMaxValue
                Me.Value = sShift
        End Select
    End If
    mTxtBtn.SetFocus
    RaiseEvent Change(mTxtBtn, mValue)
End Sub

' DblClick event handler for the slider button
' Event handler for the slider button click
' Raises the DblClick event, passing a reference to the control and the current value
' Example usage in a form:
' Private Sub slider_DblClick(control As Object, Value As Single)
'     Debug.Print "DblClick on control: " & control.Name & ", value: " & Value
' End Sub
Private Sub mTxtBtn_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If mLocked Then Exit Sub
    Cancel = True
    RaiseEvent DblClick(mTxtBtn, mValue)
End Sub

' MouseUp event handler for the slider
' Event handler for releasing the mouse button on the slider
' Updates the button position after dragging is completed
' Example usage: updating button position when mouse is released
Private Sub mTxtBtn_MouseUp(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button <> 1 Or mLocked Then Exit Sub
    With mTxtBtn
        If mIsHorizontal Then
            .Left = mLabelLineFull.Left + mLabelLineFull.Width - .Width / 2
        Else
            .Top = mLabelLineFull.Top + mLabelLineFull.Height - .Height / 2
        End If
    End With
End Sub

Private Sub mParent_MouseMove(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ApplyHoverEffect(0)
End Sub

' MouseMove event handler for the slider
' Event handler for mouse movement on the slider
' Updates the button position and slider value during dragging
' Raises the Change event when the value changes
' Example usage in a form:
' Private Sub slider_Change(control As Object, Value As Single)
'     Debug.Print "Value changed: " & Value
' End Sub
Private Sub mTxtBtn_MouseMove(ByVal Button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ApplyHoverEffect(100)
    If Button <> 1 Or mLocked Then Exit Sub
    With mTxtBtn
        If mIsHorizontal Then
            Select Case .Left + X
                Case mSliderMinPosition To mSliderMaxPosition - 5:
                    .Left = .Left + X
                    Select Case mTxtBtn.Left + X
                        Case Is < mSliderMinPosition
                            mLabelLineFull.Width = 0
                        Case Is > mSliderMaxPosition
                            mLabelLineFull.Width = mLabelLineEmpty.Width
                        Case Else
                            If mLabelLineFull.Width + X < 0 Then
                                mLabelLineFull.Width = 0
                            Else
                                mLabelLineFull.Width = mLabelLineFull.Width + X
                            End If
                    End Select
            End Select
        Else
            Select Case .Top + Y
                Case mSliderMinPosition To mSliderMaxPosition - 5:
                    .Top = .Top + Y
                    Select Case mTxtBtn.Top + Y
                        Case Is < mSliderMinPosition
                            mLabelLineFull.Height = 0
                        Case Is > mSliderMaxPosition
                            mLabelLineFull.Height = mLabelLineEmpty.Height
                        Case Else
                            If mLabelLineFull.Height + Y < 0 Then
                                mLabelLineFull.Height = 0
                            Else
                                mLabelLineFull.Height = mLabelLineFull.Height + Y
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
    RaiseEvent Change(mTxtBtn, mValue)
End Sub

' Applies a hover effect to the slider button by adjusting its color
' The shade parameter determines the intensity of the effect (0-255)
Private Sub ApplyHoverEffect(ByVal shade As Byte)
    mTxtBtn.ForeColor = BlendColor(mButtonColor, RGB(255, 255, 255), shade / 255)
End Sub

' Function to blend two colors based on a factor
' This function is used to create the hover effect on the slider button
' Parameters:
' - color1: the base color to start from
' - color2: the target color to blend towards
' - factor: the blending factor (0.0 to 1.0), where 0.0 is color1 and 1.0 is color2
Private Function BlendColor(ByVal color1 As Long, color2 As Long, factor As Double) As Long
    Dim r1 As Byte, g1 As Byte, b1 As Byte
    Dim r2 As Byte, g2 As Byte, b2 As Byte
    Dim r As Byte, g As Byte, b As Byte
    
    ' Extract RGB components from the first color (VBA uses BGR format)
    r1 = color1 Mod 256
    g1 = (color1 \ 256) Mod 256
    b1 = (color1 \ 65536) Mod 256
    
    ' Extract RGB components from the second color (VBA uses BGR format)
    r2 = color2 Mod 256
    g2 = (color2 \ 256) Mod 256
    b2 = (color2 \ 65536) Mod 256
    
    ' Blend the colors
    r = r1 + (r2 - r1) * factor
    g = g1 + (g2 - g1) * factor
    b = b1 + (b2 - b1) * factor
    
    ' Return the blended color
    BlendColor = RGB(r, g, b)
End Function

' Calculates the current slider value based on the slider position and value range
' Calculates the current slider value based on the position of the thumb and value range
' Returns a value in the range from MinValue to MaxValue
' Example usage: Dim currentValue As Single: currentValue = slider.getValue()
Private Function getValue() As Single
    ' Check for division by zero
    If mLabelLineEmpty.Width = 0 Or mLabelLineEmpty.Height = 0 Then
        getValue = mMinValue
        mValue = getValue
        Exit Function
    End If

    Dim positionRatio As Double
    If mIsHorizontal Then
        positionRatio = mLabelLineFull.Width / mLabelLineEmpty.Width
    Else
        positionRatio = mLabelLineFull.Height / mLabelLineEmpty.Height
    End If

    ' Calculate the raw value based on position ratio
    Dim rawValue As Double
    rawValue = mMinValue + positionRatio * (mMaxValue - mMinValue)

    ' Apply step increment to the value
    Dim stepValue As Double
    stepValue = (rawValue - mMinValue) / mStep
    If (stepValue - Fix(stepValue))  > 0.5 Then
        getValue = mMinValue + Fix(stepValue) * mStep + mStep
    Else
        getValue = mMinValue + Fix(stepValue) * mStep
    End If

    ' Ensure the value is within bounds
    If getValue > mMaxValue Then
        getValue = mMaxValue
    ElseIf getValue < mMinValue Then
        getValue = mMinValue
    End If

    mValue = getValue
    Call SetValue(mValue)
End Function

' Sets the slider value and updates the position of the controls
' Sets the slider value and updates the position of the controls
' Accepts a new value and adjusts the slider control positions accordingly
' Example usage: Call slider.SetValue(75)
Private Sub SetValue(ByRef snValue As Single)
    mValue = snValue
    Call UpdateSliderLinePosition
    Call UpdateButtonPosition
    Call UpdateValueLabelDisplay
End Sub

' Updates the position and size of the slider line based on the current value
Private Sub UpdateSliderLinePosition()
    With mLabelLineFull
        If mIsHorizontal Then
            .Height = mLabelLineEmpty.Height
            ' Check for division by zero
            If mMaxValue <> mMinValue Then
                .Width = (mValue - mMinValue) / (mMaxValue - mMinValue) * mLabelLineEmpty.Width
            Else
                .Width = 0
            End If
        Else
            ' Check for division by zero
            If mMaxValue <> mMinValue Then
                .Height = (mValue - mMinValue) / (mMaxValue - mMinValue) * mLabelLineEmpty.Height
            Else
                .Height = 0
            End If
            .Width = mLabelLineEmpty.Width
        End If
    End With
End Sub

' Updates the position of the slider button based on the current value and orientation
Private Sub UpdateButtonPosition()
    With mTxtBtn
        If mIsHorizontal Then
            .Top = mLabelLineEmpty.Top + mLabelLineEmpty.Height / 2 - .Height / 2 + 3
            .Left = mLabelLineFull.Left + mLabelLineFull.Width - .Width / 2
        Else
            .Top = mLabelLineFull.Top + mLabelLineFull.Height - .Height / 2
            .Left = mLabelLineEmpty.Left + mLabelLineEmpty.Width / 2 - .Width / 2 - 3
        End If
    End With
End Sub

' Updates the value label display with the current value and format
Private Sub UpdateValueLabelDisplay()
    With mLabelValue
        .Height = 10
        .Width = 30
        Call UpdateValueLabelPosition(mPositionValue)
        .Caption = mValue
        If mFormatValue <> vbNullString Then .Caption = VBA.Format$(mValue, mFormatValue)
    End With
End Sub

' Additional helper methods for improved maintainability
' These methods separate concerns and make the code more modular

' Validates that the minimum value is less than the maximum value
Private Function IsValidRange() As Boolean
    IsValidRange = (mMinValue < mMaxValue)
End Function

' Calculates the position ratio of the current value within the slider range
Private Function CalculatePositionRatio() As Double
    If mMaxValue <> mMinValue Then
        CalculatePositionRatio = (mValue - mMinValue) / (mMaxValue - mMinValue)
    Else
        CalculatePositionRatio = 0
    End If
End Function

' Ensures the given value is within the valid range
Private Function ConstrainValue(ByVal inputValue As Single) As Single
    If inputValue < mMinValue Then
        ConstrainValue = mMinValue
    ElseIf inputValue > mMaxValue Then
        ConstrainValue = mMaxValue
    Else
        ConstrainValue = inputValue
    End If
End Function

' Changes the position of the slider value label depending on the given position
' Changes the position of the slider value label based on the specified position
' Accepts a value from the PositionValue enumeration and sets the appropriate label position
' Example usage: Call slider.changePositionLabelValue(clsSlider.PositionValue.Right)
Private Sub UpdateValueLabelPosition(ByVal pos As PositionValue)
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

' Clean up resources when the class object is destroyed
' Sets all internal objects to Nothing to prevent memory leaks
Private Sub Class_Terminate()
    Set mTxtBtn = Nothing
    Set mLabelLineEmpty = Nothing
    Set mLabelLineFull = Nothing
    Set mLabelValue = Nothing
    Set mParent = Nothing
End Sub