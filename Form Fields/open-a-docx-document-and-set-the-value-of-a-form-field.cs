using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        string inputPath = @"C:\Docs\input.docx";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\output.docx";

        // Name of the form field to modify (bookmark name of the field).
        string fieldName = "MyTextInput";

        // New value to set for the text input form field.
        string newValue = "New placeholder text";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Retrieve the form field by its bookmark name.
        FormField formField = doc.Range.FormFields[fieldName];

        // Ensure the field exists and is a text input field before setting its value.
        if (formField != null && formField.Type == FieldType.FieldFormTextInput)
        {
            // Apply the new value to the form field.
            formField.SetTextInputValue(newValue);
        }

        // Update all fields in the document so the change is reflected.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save(outputPath);
    }
}
