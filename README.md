# VBA Slider Class

![Project Demo](User_Forms.gif)

This repository contains a VBA implementation of a customizable slider control that can be used in Excel UserForms. The slider class provides a more flexible and feature-rich alternative to the standard scrollbar control.

## Table of Contents
1. [Features](#features)
2. [Components](#components)
3. [Installation](#installation)
4. [Quick Start](#quick-start)
5. [Main Functions](#main-functions)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Troubleshooting](#troubleshooting)

## Features

- Customizable slider appearance (size, colors, orientation)
- Support for both horizontal and vertical orientations
- Adjustable minimum and maximum values
- Step size control for precise value selection
- Event handling for value changes
- Smooth visual updates
- Three-state functionality (unchecked, checked, indeterminate)
- Cyclic/non-cyclic state switching
- Ability to set/get state by text
- Method to get all available states
- Improved error handling and validation
- Ability to set font name and size factor
- Ability to get/set current icon
- Method to set color for specific state
- Method to reset to initial state

## Components

- `clsSlider.cls`: The main slider class implementation
- `frmTestClass.frm`: Test form demonstrating usage
- `modShowForms.bas`: Module containing form display functions
- Documentation in the `docs/` folder:
 - [`docs/technical_documentation_eng.md`](docs/technical_documentation_eng.md) - Technical documentation in English
 - [`docs/technical_documentation_rus.md`](docs/technical_documentation_rus.md) - Technical documentation in Russian
 - [`docs/user_guide_eng.md`](docs/user_guide_eng.md) - User guide in English
  - [`docs/user_guide_rus.md`](docs/user_guide_rus.md) - User guide in Russian
  - [`docs/implementation_examples_eng.md`](docs/implementation_examples_eng.md) - Implementation examples in English
 - [`docs/implementation_examples_rus.md`](docs/implementation_examples_rus.md) - Implementation examples in Russian

## Installation

1. Open the `slider_v4.xlsm` workbook in Excel
2. Import the VBA files into your project:
   - `vba-files/Class/clsSlider.cls`
   - `vba-files/Form/frmTestClass.frm`
   - `vba-files/Module/modShowForms.bas`
3. Start using the slider class in your forms

## Quick Start

### Simple Usage Example
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

## Main Functions

- **Slider Initialization**: The `Initialize` method allows setting the initial value, minimum and maximum values, and configuring appearance
- **Range Management**: Ability to configure minimum and maximum values
- **Orientation Configuration**: Support for horizontal and vertical orientations
- **Event Handling**: Support for value change and click events
- **Smooth Updates**: Visual representation of value changes

## Working with Controls

The `clsSlider` class uses a Label control as the slider track with capabilities:
- Setting initial value
- Configuring minimum and maximum values
- Handling drag events
- Displaying value label with customizable positions

## Style Configuration

The class allows customization of:
- Track colors (empty/filled portions)
- Slider button colors
- Value label colors
- Background colors
- Fonts and text sizes

Example of color configuration:
```vba
' Configure colors during initialization
Slider.Initialize Me.Label1, 50, 0, 100, True, , RGB(200, 200, 200), RGB(0, 100, 200)
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

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.