using System;
using Aspose.Words;
using Aspose.Words.Fields;

class SetFormFieldValue
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\OutputDocument.docx";

        // Load the existing document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve the form field by its bookmark name.
        // Replace "MyTextInput" with the actual name of the form field you want to modify.
        FormField formField = doc.Range.FormFields["MyTextInput"];

        if (formField != null)
        {
            // Check that the field is a text input field before setting its value.
            if (formField.Type == FieldType.FieldFormTextInput)
            {
                // Set the new value for the text input form field.
                formField.SetTextInputValue("New value for the field");
            }
            else
            {
                // For non‑text fields you can set the Result property directly.
                formField.Result = "New value for non‑text field";
            }
        }
        else
        {
            Console.WriteLine("Form field 'MyTextInput' was not found in the document.");
        }

        // Save the modified document to the specified output path.
        doc.Save(outputPath);
    }
}
