# VBAForm2HTML - VBA UserForm to HTML/CSS Converter
This program converts userforms created in Microsoft Office VBA into HTML/CSS code.<br>

## System Requirements
- Supported OS: Windows XP or later
- Required Software: Microsoft Excel/Word/PowerPoint/Outlook 2000 or later
- Recommended Environment: Microsoft Excel 2016 or later

## Verified Operating Environments
- Windows XP(SP3)
- Windows 10/11
- Excel 2000(32bit)
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)
- Word/PowerPoint/Outlook 2000 (32bit)
- Word/PowerPoint/Outlook 2019 (64bit)

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
- Images Embedded in Controls (`Image`)
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
>
>`SpinButton` behaves differently in VBA and CSS, so appearance may vary depending on placement.<br>

If unsupported controls exist on the form, the conversion will fail. If that case, please remove those controls and run the conversion again.<br>

## Usage
In the Immediate Window, enter: `Call ConvertForm2HTML(UserForm1)`<br>
```vb
Call ConvertForm2HTML(UserForm1)
```
   > Note: Replace `UserForm1` with the object name of the form you want to convert.

If conversion succeeds, a message will appear, and an `output.html` file will be created.<br>

## Output Directory

A dedicated `VBAForm2HTML_output` folder is automatically created in the workbook directory, and all generated files are saved there:

### Excel and Word

When running from Excel or Word, the output folder is created in the same directory as the macro-enabled document.

-   **Excel**: Uses the workbook's directory (`ThisWorkbook.Path`)
-   **Word**: Uses the document's directory (`MacroContainer.Path`)

```
WorkbookFolder/
├─ MyWorkbook.xlsm
└─ VBAForm2HTML_output/
   ├─ output.html
   └─ exported images...
```

### Other Office Applications
When running from other Office applications (such as PowerPoint, Outlook, etc.), or when the current Excel workbook or Word document has not yet been saved, the output folder is created in the user's **Documents** folder instead.

```
C:\Users\%USERNAME%\Documents\
└─ VBAForm2HTML_output/
   ├─ output.html
   └─ exported images...
```

If the Documents folder cannot be resolved, the output folder will be created in the root of the **C:** drive as a final fallback.

## Parameters

`ConvertForm2HTML` accepts the following parameters:

|**Parameter**|**Type**|**Description**                         |
|----------------|-------------------------------|-----------------------------|
|`frms` |`Variant`|**Required.**<br>Accepts a single `UserForm` object or an `Array` of `UserForm` objects to be converted.            |
|`usePrefix`  |`Boolean` |**Optional (Default: `False`).**<br>If set to `True`, the form name will be added to each element name. This is automatically set to `True` if `frms` is an array.|
|`imageMode`  |`String` |**Optional (Default: `"file"`).**<br>Determines how image files used in the UserForm are handled during conversion. You can choose one of the following options:<br>• `"file"` (Default): Images are saved as separate external files in the output directory, and the generated code references these files.<br>• `"disabled"`: Image processing is disabled, and no image-related code is generated.<br>• `"reference-only"`: Similar to `"file"`, generates code that references image files, but does not export the image files. Useful when the image files already exist.<br>• `"base64"`: Images are embedded directly into the generated code as Base64-encoded strings, keeping everything in a single file.|
|`langAttribute`  |`String` |**Optional (Default: `""`).**<br>Specifies the language code for the `"lang"` attribute of the generated `<html>` tag (e.g., `"en"` for English, `"zh"` for Chinese, `"ja"` for Japanese).<br>If omitted, the attribute will be left empty.|

You can execute the conversion by calling the `ConvertForm2HTML` with a single UserForm object or an array of multiple UserForms.

```vb
' Example: Converting a single form
Call ConvertForm2HTML(UserForm1)

' Example: Converting multiple forms
Call ConvertForm2HTML(Array(UserForm1, UserForm2))

' Example: Converting a single form (With image streams embedded directly as Base64 text)
Call ConvertForm2HTML(UserForm1, imageMode:="base64")
```

## Control Order (for Controls Without Child Elements)
In HTML/CSS, if you place one `<div>` on top of another, the later element appears in front.<br>
However, in VBA, you can change front/back order, so the behavior differs.<br>
The program first sorts controls by hierarchy level; however, it preserves the original creation order within the same hierarchy.<br>
Since VBA’s z-order (front/back) cannot currently be retrieved, some displays may not match VBA.<br>
