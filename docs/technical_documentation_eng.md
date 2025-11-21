# Technical Documentation for clsSlider

## Table of Contents
1. [Class Overview](#class-overview)
2. [Class Architecture](#class-architecture)
3. [Properties](#properties)
4. [Methods](#methods)
5. [Events](#events)
6. [Constants and Enumerations](#constants-and-enumerations)
7. [Implementation Details](#implementation-details)
8. [Dependencies](#dependencies)

## Class Overview

The `clsSlider` class implements a custom slider control for VBA UserForms. The slider allows users to select a value within a specified range by dragging a button along a track. It supports both horizontal and vertical orientations and provides customizable appearance and behavior options.

### Key Features
- Horizontal and vertical orientation support
- Configurable min/max values and initial value
- Customizable appearance (colors, icons, formatting)
- Visible/hidden value label with positioning options
- Event handling for value changes and clicks
- Lock functionality to prevent user interaction
- Automatic positioning of value label (top, left, right, bottom)

## Class Architecture

### PositionValue Enumeration
Defines possible positions for the slider value label:
- `Top` (1) - value label appears above the slider
- `Left` - value label appears to the left of the slider
- `Right` - value label appears to the right of the slider
- `Bottom` - value label appears below the slider

### Core Class Properties

| Property | Type | Description |
|----------|------|-------------|
| `Value` | Single | Current slider value |
| `MinValue` | Single | Minimum slider value |
| `MaxValue` | Single | Maximum slider value |
| `IsHorizontal` | Boolean | Flag for horizontal orientation |
| `Locked` | Boolean | Flag for locking the slider |
| `Visible` | Boolean | Visibility state of all slider elements |
| `Enabled` | Boolean | Enabled state of the slider |
| `VisibleLabelValue` | Boolean | Visibility state of the slider value label |
| `ForeColorValue` | XlRgbColor | Text color of the slider value label |
| `ForeColorBtn` | XlRgbColor | Text color of the slider button |
| `BackColorFull` | XlRgbColor | Color of the filled part of the slider line |
| `BackColorEmpty` | XlRgbColor | Color of the empty part of the slider line |
| `PositionLabelValue` | PositionValue | Position of the slider value label |
| `FormatValue` | String | Format for displaying the slider value |
| `ForeColorValue` | XlRgbColor | Text color of the slider value label |
| `ForeColorBtn` | XlRgbColor | Text color of the slider button |
| `Name` | String | Name of the main slider label |
| `Icon` | Long | Icon on the slider button (as a Unicode character value) |

## Properties

### Core Properties
- `Value` - Gets or sets the current value of the slider
- `MinValue` - Gets or sets the minimum value of the slider
- `MaxValue` - Gets or sets the maximum value of the slider
- `IsHorizontal` - Gets whether the slider is horizontal
- `Locked` - Gets or sets the lock state of the slider
- `Visible` - Gets or sets the visibility state of all slider elements
- `Enabled` - Gets or sets the enabled state of the slider
- `VisibleLabelValue` - Gets or sets the visibility state of the slider value label
- `ForeColorValue` - Gets or sets the text color of the slider value label
- `ForeColorBtn` - Gets or sets the text color of the slider button
- `BackColorFull` - Gets or sets the color of the filled part of the slider line
- `BackColorEmpty` - Gets or sets the color of the empty part of the slider line
- `PositionLabelValue` - Gets or sets the position of the slider value label
- `FormatValue` - Gets or sets the format for displaying the slider value
- `Name` - Gets the name of the main slider label
- `Icon` - Gets or sets the icon on the slider button

### Additional Properties
- `Version` - Gets version information about the class
- `Items` - Gets the collection of all slider items
- `Count` - Gets the number of items in the collection

## Methods

### Initialize
Initializes the slider with specified parameters.

**Syntax:**
```vba
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
        Optional Icon As Long = 5963)
```

**Parameters:**
- `labelSlider` - the main label that will be used as the empty slider line
- `Value` - initial value of the slider
- `MinValue` - minimum value of the slider
- `MaxValue` - maximum value of the slider
- `VisibleLabelValue` - visibility of the value label
- `FormatValue` - format for displaying the value (default vbNullString)
- `PositionValue` - position of the value label (default Top)
- `BackColorEmpty` - color of the empty part of the line (default rgbLightGray)
- `BackColorFull` - color of the filled part of the line (default rgbGreenYellow)
- `ForeColorBtn` - color of the button (default rgbBlack)
- `ForeColorValue` - color of the value (default rgbBlack)
- `Icon` - icon on the button (default 59963)

## Events

### Click
Event raised when the slider button is clicked.

**Syntax:**
```vba
Private Sub mLabelBtn_Click()
```

### Change
Event raised when the slider value changes.

**Syntax:**
```vba
Private Sub mLabelBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
```

## Constants and Enumerations

### PositionValue
```vba
Public Enum PositionValue
    Top = 1  ' Position of the value label above the slider
    Left   ' Position of the value label to the left of the slider
    Right ' Position of the value label to the right of the slider
    Bottom ' Position of the value label below the slider
End Enum
```

## Implementation Details

### Orientation Support
The class supports both horizontal and vertical orientations:
- Horizontal: The slider moves left and right
- Vertical: The slider moves up and down

### Visual Elements
The slider consists of multiple visual elements:
- Main slider line (empty part)
- Filled slider line (filled part)
- Slider button (movable element)
- Value label (displays current value)

### Event Handling
The class automatically handles slider events:
- `Click` - raises the Click event when the button is clicked
- `MouseMove` - updates the slider position and value when dragging

## Dependencies

- MSForms library for Label controls
- XlRgbColor constants for color properties