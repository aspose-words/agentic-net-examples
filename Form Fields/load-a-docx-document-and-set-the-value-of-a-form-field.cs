using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Name of the form field to modify (bookmark name of the field).
        string fieldName = "MyTextInput";

        // New value to assign to the form field.
        string newValue = "New value";

        // Load the existing document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve the form field by its bookmark name.
        FormField formField = doc.Range.FormFields[fieldName];

        if (formField != null)
        {
            // If the field is a text input, use SetTextInputValue to apply formatting.
            if (formField.Type == FieldType.FieldFormTextInput)
            {
                formField.SetTextInputValue(newValue);
            }
            else
            {
                // For other field types (e.g., checkbox, dropdown) set the result directly.
                formField.Result = newValue;
            }
        }
        else
        {
            Console.WriteLine($"Form field \"{fieldName}\" not found.");
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
