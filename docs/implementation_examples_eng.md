# Implementation Examples for clsSlider

## Table of Contents
1. [Basic Implementation](#basic-implementation)
2. [Advanced Configuration](#advanced-configuration)
3. [Custom Styling](#custom-styling)
4. [Event Handling](#event-handling)
5. [Integration with Forms](#integration-with-forms)
6. [Dynamic Slider Creation](#dynamic-slider-creation)
7. [Working with Multiple Sliders](#working-with-multiple-sliders)
8. [Advanced Examples](#advanced-examples)

## Basic Implementation

### Simple Slider Usage
```vba
' In the form module
Dim Slider As clsSlider

Private Sub UserForm_Initialize()
    ' Create an instance of the slider class
    Set Slider = New clsSlider
    
    ' Initialize the slider with default parameters
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub
```

### Slider with Value Display
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Initialize with visible value label
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, "0.0")
    
    ' Update a label with current value
    LabelValue.Caption = Slider.Value
End Sub

Private Sub Slider_ValueChanged()
    ' Update when slider value changes
    LabelValue.Caption = Slider.Value
End Sub
```

## Advanced Configuration

### Configuring Colors and Appearance
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Initialize with custom colors
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, "0.0", clsSlider.PositionValue.Right, _
        RGB(200, 200, 200), RGB(0, 100, 200), RGB(0, 0, 0), RGB(0, 0, 0))
        
    ' Further customize appearance
    With Slider
        .ForeColorValue = RGB(0, 100, 200)
        .BackColorFull = RGB(0, 150, 255)
        .BackColorEmpty = RGB(220, 220, 220)
    End With
End Sub
```

### Horizontal vs Vertical Slider
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Horizontal slider (default) - wider than tall
    Call Slider.Initialize(Me.HorizontalLabel, 50, 0, 100, True)
    
    ' Vertical slider - taller than wide
    Call Slider.Initialize(Me.VerticalLabel, 50, 0, 100, True)
    
    ' Check orientation
    If Slider.IsHorizontal Then
        Debug.Print "Slider is horizontal"
    Else
        Debug.Print "Slider is vertical"
    End If
End Sub
```

## Custom Styling

### Custom Value Format
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Initialize with custom value formatting
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, "0.00%", clsSlider.PositionValue.Top)
    
    ' This will display values as percentages (e.g., 50%)
End Sub
```

### Slider with Icon
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    
    ' Initialize with custom icon on the button
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, , , , , , , 59963) ' Using icon code
End Sub
```

## Event Handling

### Handling Value Changes
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub

Private Sub mLabelBtn_Click()
    ' Handle click event
    Debug.Print "Slider clicked, current value: " & Slider.Value
End Sub

Private Sub mLabelBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Handle mouse move (dragging)
    Debug.Print "Slider value changed to: " & Slider.Value
    ' Update other controls based on slider value
    TextBox1.Value = Slider.Value
End Sub
```

### Slider Locking
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
    
    ' Initially lock the slider
    Slider.Locked = True
End Sub

Private Sub LockSliderButton_Click()
    ' Toggle slider lock state
    Slider.Locked = Not Slider.Locked
    If Slider.Locked Then
        LockButton.Caption = "Unlock Slider"
    Else
        LockButton.Caption = "Lock Slider"
    End If
End Sub
```

## Integration with Forms

### Complete Form Example
```vba
Dim VolumeSlider As clsSlider
Dim BrightnessSlider As clsSlider

Private Sub UserForm_Initialize()
    ' Initialize volume slider
    Set VolumeSlider = New clsSlider
    Call VolumeSlider.Initialize(Me.VolumeLabel, 75, 0, 100, True, "0", clsSlider.PositionValue.Right)
    
    ' Initialize brightness slider
    Set BrightnessSlider = New clsSlider
    Call BrightnessSlider.Initialize(Me.BrightnessLabel, 50, 0, 100, True, "0", clsSlider.PositionValue.Right)
    
    ' Update form controls with initial values
    UpdateVolumeDisplay
    UpdateBrightnessDisplay
End Sub

Private Sub UpdateVolumeDisplay()
    VolumeValueLabel.Caption = "Volume: " & VolumeSlider.Value & "%"
    ' Simulate volume change effect
    ' (In real application, this would control audio volume)
End Sub

Private Sub UpdateBrightnessDisplay()
    BrightnessValueLabel.Caption = "Brightness: " & BrightnessSlider.Value & "%"
    ' Simulate brightness change effect
    ' (In real application, this would control screen brightness)
End Sub

Private Sub VolumeLabel_ValueChanged()
    UpdateVolumeDisplay
End Sub

Private Sub BrightnessLabel_ValueChanged()
    UpdateBrightnessDisplay
End Sub
```

## Dynamic Slider Creation

### Creating Sliders Programmatically
```vba
Private Sub UserForm_Initialize()
    ' Create multiple sliders dynamically
    Dim i As Integer
    Dim newLabel As MSForms.Label
    Dim newSlider As clsSlider
    
    For i = 1 To 3
        ' Create a new label for the slider
        Set newLabel = Me.Controls.Add("Forms.Label.1", "DynamicSlider" & i, True)
        With newLabel
            .Left = 20
            .Top = 30 + (i - 1) * 40
            .Width = 200
            .Height = 5
            .BackColor = RGB(200, 200, 200)
        End With
        
        ' Create and initialize slider
        Set newSlider = New clsSlider
        Call newSlider.Initialize(newLabel, 50, 0, 100, True, , clsSlider.PositionValue.Bottom)
        
        ' Store reference in a collection for later access
        SliderCollection.Add newSlider
    Next i
End Sub
```

## Working with Multiple Sliders

### Managing Multiple Sliders
```vba
Dim SliderCollection As Collection

Private Sub UserForm_Initialize()
    Set SliderCollection = New Collection
    
    ' Initialize multiple sliders
    InitializeSliders
End Sub

Private Sub InitializeSliders()
    Dim RedSlider As clsSlider
    Dim GreenSlider As clsSlider
    Dim BlueSlider As clsSlider
    
    ' Create RGB sliders
    Set RedSlider = New clsSlider
    Set GreenSlider = New clsSlider
    Set BlueSlider = New clsSlider
    
    ' Initialize with different parameters
    Call RedSlider.Initialize(Me.RedLabel, 128, 0, 255, True, "0", clsSlider.PositionValue.Right)
    Call GreenSlider.Initialize(Me.GreenLabel, 128, 0, 255, True, "0", clsSlider.PositionValue.Right)
    Call BlueSlider.Initialize(Me.BlueLabel, 128, 0, 255, True, "0", clsSlider.PositionValue.Right)
    
    ' Add to collection
    SliderCollection.Add RedSlider
    SliderCollection.Add GreenSlider
    SliderCollection.Add BlueSlider
    
    ' Update color preview
    UpdateColorPreview
End Sub

Private Sub UpdateColorPreview()
    ' Combine slider values to create color
    Dim combinedColor As Long
    combinedColor = RGB(SliderCollection(1).Value, SliderCollection(2).Value, SliderCollection(3).Value)
    
    ' Apply color to preview area
    ColorPreview.BackColor = combinedColor
    ColorPreview.Caption = "RGB(" & SliderCollection(1).Value & "," & _
                           SliderCollection(2).Value & "," & _
                           SliderCollection(3).Value & ")"
End Sub

Private Sub RedLabel_ValueChanged()
    UpdateColorPreview
End Sub

Private Sub GreenLabel_ValueChanged()
    UpdateColorPreview
End Sub

Private Sub BlueLabel_ValueChanged()
    UpdateColorPreview
End Sub
```

## Advanced Examples

### Slider with Animation Effect
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub

Private Sub AnimateSliderEffect()
    Dim originalColor As Long
    originalColor = Slider.BackColorFull
    
    ' Create animation effect by changing color temporarily
    Slider.BackColorFull = RGB(255, 255, 0) ' Yellow
    
    ' Wait briefly to show effect
    Application.Wait (Now + TimeValue("0:00:0.2"))
    
    ' Restore original color
    Slider.BackColorFull = originalColor
End Sub

Private Sub mLabelBtn_Click()
    ' Trigger animation on click
    AnimateSliderEffect
End Sub
```

### Responsive Slider Panel
```vba
Private Sub UserForm_Resize()
    ' Adjust slider positions when form resizes
    AdjustSliderLayout
End Sub

Private Sub AdjustSliderLayout()
    If Not Slider Is Nothing Then
        With Slider
            ' Adjust slider to maintain aspect ratio
            .control.Width = Me.Width - 60
            .control.Left = 30
            
            ' If slider is vertical, adjust height
            If Not .IsHorizontal Then
                .control.Height = Me.Height - 100
                .control.Top = 50
            End If
        End With
    End If
End Sub
```

These examples demonstrate various ways to implement and use the `clsSlider` class in VBA UserForms. Each example can be adapted to specific application requirements.