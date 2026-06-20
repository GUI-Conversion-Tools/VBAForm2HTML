# VBAForm2HTML - Excel VBA UserForm to HTML/CSS Converter
This program converts userforms created in Microsoft Excel VBA into HTML/CSS code.<br>

## System Requirements
- Supported OS: Windows
- Required Software: Microsoft Excel 2000 or later
- Recommended Environment: Microsoft Excel 2016 or later

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Control colors (foreground)
- Control colors (background) (Excluding `MultiPage`, `ScrollBar`)
- Text display (`UserForm`, `Frame`, `Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- Font (typeface, size, bold, italic)
- Borders (`UserForm`, `Frame`, `TextBox`, `Label`, `ListBox`, `ComboBox`, `Image`)
- Mouse cursor
- Text alignment: left, center, right (`Label`, `TextBox`, `ComboBox`, `CheckBox`, `ToggleButton`, `OptionButton`, `ListBox`)
- Default values of `TextBox`, `ComboBox`
- Items set in `ComboBox`, `ListBox`
- Selection state of `OptionButton`, `CheckBox`, `ToggleButton`
- Transparent background setting specified in `.BackStyle`(`Label`, `TextBox`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `Image`, `ComboBox`)
- `.TabOrientation` property (`MultiPage`)
- `.Locked` property (`TextBox`)
- `.PasswordChar` property (`TextBox` [.MultiLine=False])
- `.ScrollBars` property (`TextBox` [.MultiLine=True])
- `.WordWrap` property (`TextBox` [.MultiLine=True])
- `.Style` property (`ComboBox`, `MultiPage`)
- `.MultiSelect` property (`ListBox`)
- `.PictureAlignment` property (`Image`)

> Notes:
>
>-   For `ListBox`, `.MultiSelect` values `fmMultiSelectMulti` and `fmMultiSelectExtended` are both converted to HTML multiple-selection mode (`multiple`). After conversion, selecting multiple items follows standard browser behavior and typically requires holding `Ctrl` (or the platform equivalent modifier key).
>-   For `TextBox`, `.PasswordChar` is reflected only by converting the control to a password input (`type="password"`); the actual masking character itself is not preserved.
>-   For `MultiPage`, `.TabOrientation` supports `fmTabOrientationTop` and `fmTabOrientationBottom` only. Any other value is treated as `fmTabOrientationTop`.
>-   For `MultiPage`, `.Style` values `fmTabStyleTabs` and `fmTabStyleButtons` are rendered identically in the generated HTML. (`fmTabStyleNone` remains handled separately.)


## Supported Controls
| VBA Form Class | CSS Element|
| ------ | ------ |
| `Label` | `<div>` |
| `CommandButton` | `<button>` |
| `Frame` | `<fieldset>` + `<legend>` |
| `TextBox` (`.MultiLine=False`) | `<input>` |
| `TextBox` (`.MultiLine=True`) | `<textarea>` |
| `SpinButton` | `<input type="number">` |
| `ListBox` | `<select size=[n]>` |
| `CheckBox` | `<label>` + `<input type="checkbox">` |
| `ToggleButton` | `<input type="checkbox">` (`display:none`) + `<label>` |
| `OptionButton` | `<label>` + `<input type="radio">` |
| `Image` | `<div>` |
| `ScrollBar` | `<input type="range">` |
| `ComboBox` (`.Style = fmStyleDropDownCombo`) | `<input list="[datalist name]">` |
| `ComboBox` (`.Style = fmStyleDropDownList`) | `<select>` |
| `MultiPage` | `<div>` (Tabs) + `<div>` (Pages) |

> Note:
`SpinButton` behaves differently in VBA and CSS, so appearance may vary depending on placement.<br>
If unsupported controls exist on the form, the conversion will fail. If that case, please remove those controls and run the conversion again.<br>

## Usage
In the Immediate Window, enter: `Call ConvertForm2HTML(UserForm1)`<br>
```vb
Call ConvertForm2HTML(UserForm1)
```
   > Note: Replace `UserForm1` with the object name of the form you want to convert.

If conversion succeeds, a message will appear, and an `output.html` file will be created in the same directory as your Excel workbook.<br>

## Parameters

`ConvertForm2HTML` accepts the following parameters:

|**Parameter**|**Type**|**Description**                         |
|----------------|-------------------------------|-----------------------------|
|`frms` |`Variant`|**Required.**<br>Accepts a single `UserForm` object or an `Array` of `UserForm` objects to be converted.            |
|`usePrefix`  |`Boolean` |**Optional (Default: `False`).**<br>If set to `True`, the form name will be added to each element name. This is automatically set to `True` if `frms` is an array.|

You can execute the conversion by calling the `ConvertForm2HTML` with a single UserForm object or an array of multiple UserForms.

```vb
' Example: Converting a single form
Call ConvertForm2HTML(UserForm1)

' Example: Converting multiple forms
Call ConvertForm2HTML(Array(UserForm1, UserForm2))
```

## Control Order (for Controls Without Child Elements)
In HTML/CSS, if you place one `<div>` on top of another, the later element appears in front.<br>
However, in VBA, you can change front/back order, so the behavior differs.<br>
The program first sorts controls by hierarchy level; however, it preserves the original creation order within the same hierarchy.<br>
Since VBA’s z-order (front/back) cannot currently be retrieved, some displays may not match VBA.<br>
