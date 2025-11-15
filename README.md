![User Forms Slider Demo](User_Forms.gif)

# VBA Slider Class

This repository contains a VBA implementation of a customizable slider control that can be used in Excel UserForms. The slider class provides a more flexible and feature-rich alternative to the standard scrollbar control.

## Features

- Customizable slider appearance (size, colors, orientation)
- Support for both horizontal and vertical orientations
- Adjustable minimum and maximum values
- Step size control for precise value selection
- Event handling for value changes
- Smooth visual updates

## Components

- `clsSlider.cls`: The main slider class implementation
- `frmTestClass.frm`: Test form demonstrating slider usage
- `modShowForms.bas`: Module containing form display functions
- Documentation in the `docs/` folder

## Installation

1. Open the `slider_v2.xlsm` workbook in Excel
2. Import the VBA files into your project:
   - `vba-files/Class/clsSlider.cls`
   - `vba-files/Form/frmTestClass.frm`
   - `vba-files/Module/modShowForms.bas`
3. Start using the slider class in your forms

## Usage

To use the slider in your UserForm:

1. Create an instance of the clsSlider class
2. Initialize it with your UserForm and desired parameters
3. Set properties like min/max values, step size, and appearance
4. Handle the ValueChanged event to respond to user interactions

See the `frmTestClass` form for a complete example implementation.

## Documentation

Detailed documentation for the slider class is available in `docs/clsSlider_Documentation.md`.

## License

This project is licensed under the terms found in the LICENSE file.