# SearchTBox

Finds, highlights, and counts all instances of a matching string within text boxes.

## Compatibility:

Microsoft Excel 2016

## Installation:

1.) Open the workbook that contains text boxes you'd like to search in. If you do not have the **Developer** tab active in your worksheet,
please follow this [tutorial](https://support.microsoft.com/en-us/office/add-or-edit-a-macro-for-a-control-on-a-worksheet-3b1f0fa9-e988-40e0-8f5f-40fe9c1f7126) under **'Add or edit a macro for an ActiveX control'.**

2.) Navigate to **Developer** > **Visual Basic.**

3.) Right click on 'ThisWorkbook' in the left menu bar, and select 'Import File.'

4.) Import the 'SearchTBox.bas' file and exit the Visual Basic window.

5.) In your Excel worksheet, navigate to **Developer** > **Macros**.

6.) Enter 'SearchTBox' as the 'Macro name' and click 'Run.'

7.) Enter the string you'd like to search for. On a match, the text box will turn black and the text color will turn white. A pop-up message box also returns the number of matches found.
