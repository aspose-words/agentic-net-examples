using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractFormFieldContent
{
    static void Main()
    {
        // Load the WORDML (DOCX) document.
        Document doc = new Document("Input.docx");

        // Update all fields so that the latest values are available.
        doc.Range.UpdateFields();

        // Store extracted values: key = form field name, value = displayed content.
        var fieldValues = new Dictionary<string, string>();

        // Iterate through every form field in the document.
        foreach (FormField formField in doc.Range.FormFields)
        {
            string value = GetFormFieldValue(formField);
            fieldValues[formField.Name] = value;
        }

        // Output the extracted values.
        foreach (var kvp in fieldValues)
        {
            Console.WriteLine($"{kvp.Key}: {kvp.Value}");
        }
    }

    // Returns the displayed content of a given form field.
    private static string GetFormFieldValue(FormField field)
    {
        switch (field.Type)
        {
            case FieldType.FieldFormTextInput:
                // Text input field – the current text entered by the user.
                return field.Result;

            case FieldType.FieldFormCheckBox:
                // Checkbox – indicate whether it is checked.
                return field.Checked ? "Checked" : "Unchecked";

            case FieldType.FieldFormDropDown:
                // Drop‑down list – return the selected item text.
                int index = field.DropDownSelectedIndex;
                if (index >= 0 && index < field.DropDownItems.Count)
                    return field.DropDownItems[index];
                return string.Empty;

            default:
                return string.Empty;
        }
    }
}
