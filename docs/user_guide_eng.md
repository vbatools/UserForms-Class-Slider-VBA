# User Guide for clsSlider

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Installation and Setup](#installation-and-setup)
4. [Quick Start](#quick-start)
5. [Main Features](#main-features)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Creating Toggle Switches](#creating-toggle-switches)
9. [Adding Icons](#adding-icons)
10. [Working with Style Collection](#working-with-style-collection)
11. [Troubleshooting](#troubleshooting)
12. [Frequently Asked Questions](#frequently-asked-questions)

## Introduction

The `clsSlider` class implements a custom slider control for VBA UserForms. The slider allows users to select a value within a specified range by dragging a button along a track. It supports both horizontal and vertical orientations and provides customizable appearance and behavior options.

### What This Project Can Do:
- Apply modern design to slider controls
- Provide smooth value selection with visual feedback
- Configure colors and fonts for sliders
- Add icons and visual elements
- Manage visibility and state of slider controls

## System Requirements

- Microsoft Excel (2010 or newer recommended)
- VBA support enabled
- Microsoft Forms 2.0 Object Library
- Windows 7 or newer

## Installation and Setup

### Step 1: Import the Class
1. Open Excel and go to the VBA editor (press Alt+F11)
2. In the menu, select "File" > "Import File"
3. Select the `clsSlider.cls` file from the `vba-files/Class/` directory
4. Click "Open" to import the class

### Step 2: Configure References
1. In the VBA editor, select "Tools" > "References"
2. Find and check the box next to "Microsoft Forms 2.0 Object Library"
3. Click "OK" to save changes

### Step 3: Create a User Form
1. In the VBA editor, create a new user form
2. Add a Label control that will serve as the slider track
3. Add a class variable to the form:
```vba
Dim Slider As clsSlider
```

## Quick Start

### Simple Usage Example
1. Create a new user form in Excel
2. Add a Label control to serve as the slider track
3. In the `UserForm_Initialize` event, add the following code:
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True)
End Sub
```
4. Run the form to see the slider control

### Example with Color Configuration
```vba
Private Sub UserForm_Initialize()
    Set Slider = New clsSlider
    Call Slider.Initialize(Me.Label1, 50, 0, 100, True, , , RGB(200, 200, 200), RGB(0, 100, 200), RGB(0, 0, 0), RGB(0, 0, 0))
End Sub
```

## Main Features

### Slider Initialization
The `Initialize` method is the main way to set up the slider:
- Sets the initial value, minimum and maximum values
- Configures appearance properties
- Sets up event handlers for interaction
- Creates value label if enabled

### Supported Orientations
- Horizontal - slider moves left and right
- Vertical - slider moves up and down
- Automatic detection based on control dimensions

## Working with Controls

### Slider Configuration
For sliders, the class automatically:
- Applies the specified value range
- Creates a movable button element
- Shows a value label with positioning options
- Provides visual feedback during interaction
- Handles mouse events for dragging

### Value Range
The class allows configuring:
- Minimum and maximum values
- Initial value
- Value display format
- Position of the value label

### Event Handling
The class provides events for:
- Value changes during interaction
- Click events on the slider button
- Visual feedback during dragging

## Style Configuration

### Color Schemes
The class provides extensive color configuration options:
- Track colors (empty/filled portions)
- Button colors
- Value label colors
- Background colors

### Configuration via Initialize Method
Colors can be configured during initialization:
```vba
Slider.Initialize Me.Label1, 50, 0, 100, True, , , RGB(200, 200, 200), RGB(0, 100, 200)
```

### Dynamic Configuration
After initialization, you can change properties:
```vba
With Slider
    .BackColorFull = RGB(255, 0, 0)
    .ForeColorValue = RGB(0, 0, 255)
End With
```

### Font Configuration
The class allows:
- Setting font size for value labels
- Configuring font properties via `FormatValue` property

## Creating Toggle Switches

The `clsSlider` class does not implement toggle switches directly. For toggle switches, consider using the `clsMultiStateCheckBox` class with the SWITCH tag property.

## Adding Icons

### Using Built-in Icons
The slider class uses icons primarily for the slider button:
- Customizable button appearance
- Unicode character support
- Font configuration for icons

### Setting Icons
To customize the slider button:
1. In the Initialize method, specify the Icon parameter
2. Or programmatically:
```vba
Slider.Icon = 59963  ' Using numeric value for icon
```

### Icon Configuration
- Icons are displayed using the Segoe MDL2 Assets font by default
- Icon color changes depending on the slider state
- Icons automatically scale to the button size

## Working with Style Collection

### Accessing Individual Elements
After initialization, the slider provides access to its elements:
```vba
' Getting slider properties
Dim currentValue As Single
currentValue = Slider.Value
```

### Getting the Number of Elements
The slider class manages a single control element, but provides access to its properties.

### Iterating Through All Elements
The slider class focuses on a single control, so iteration is not typically necessary.

### Changing Properties of Individual Elements
```vba
' Changing slider properties
With Slider
    .Value = 75
    .MinValue = 0
    .MaxValue = 100
End With
```

## Troubleshooting

### Display Issues
- Ensure Microsoft Forms 2.0 Object Library is enabled in references
- Check that the Label control is added before calling the Initialize method
- Ensure the MultiUse property is set to True for the class

### Interaction Issues
- Check that control events are not overloaded with other handlers
- Ensure control properties are not changed manually while the class is running
- Verify that the class is not initialized multiple times

### Performance Issues
- Reduce the complexity of value formatting
- Avoid frequent calls to property getters during interaction
- Use visibility and availability properties appropriately

### Common Errors
- "Object variable not set" - ensure the class variable is properly initialized
- "Method or data member not found" - check that the class is properly imported
- "Can't assign to property" - avoid direct assignment to nested objects without checking for Nothing

## Frequently Asked Questions

### Question: How to change colors after initialization?
**Answer:** Use the properties of the slider instance to change colors after initialization.

### Question: Are all control types supported?
**Answer:** The class specifically supports slider functionality using Label controls as tracks.

### Question: Can multiple class instances be used?
**Answer:** Yes, you can create multiple class instances for different sliders, but each instance manages a single slider control.

### Question: How to add custom icons?
**Answer:** The class uses Unicode character values for icons. For custom icons, you can use different character codes or consider using image controls.

### Question: Is the class compatible with different Excel versions?
**Answer:** The class is tested with Excel 2010 and newer. Compatibility with earlier versions is not guaranteed.

### Question: Can animation be configured?
**Answer:** The current version provides smooth value transition during dragging, but direct animation configuration is not available.

### Question: How to handle events of slider elements?
**Answer:** The class provides Click and value change events that can be handled in the form module.