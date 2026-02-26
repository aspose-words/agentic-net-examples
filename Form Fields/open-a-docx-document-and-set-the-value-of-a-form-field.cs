using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the existing DOCX file.
        const string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the modified document will be saved.
        const string outputPath = @"C:\Docs\OutputDocument.docx";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Retrieve the form field by its bookmark name.
        // Replace "MyTextInput" with the actual name of the form field you want to modify.
        FormField formField = doc.Range.FormFields["MyTextInput"];

        // Ensure the field exists and is a text input field before setting its value.
        if (formField != null && formField.Type == FieldType.FieldFormTextInput)
        {
            // Set the new value for the text input form field.
            formField.SetTextInputValue("New value for the form field");
        }

        // Save the updated document.
        doc.Save(outputPath);
    }
}
