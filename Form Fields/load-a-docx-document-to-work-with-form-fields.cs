using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the DOCX file that contains form fields.
        string inputPath = @"C:\Docs\SampleForm.docx";

        // Load the document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            Console.WriteLine($"Field Name: {field.Name}, Type: {field.Type}");

            // Example: set a new value for text input fields.
            if (field.Type == FieldType.FieldFormTextInput)
            {
                field.Result = "New value";
            }
        }

        // Save the modified document (optional).
        string outputPath = @"C:\Docs\SampleForm_Updated.docx";
        doc.Save(outputPath);
    }
}
