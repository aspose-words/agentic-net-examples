using System;
using Aspose.Words;
using Aspose.Words.Fields;

class LoadDocumentExample
{
    static void Main()
    {
        // Path to the DOCX file that contains form fields.
        string docPath = @"C:\Docs\SampleForm.docx";

        // Load the existing document from the file system.
        Document doc = new Document(docPath);

        // Iterate through all form fields in the document.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Example: Print the name and current value of each form field.
            Console.WriteLine($"Field Name: {field.Name}");
            Console.WriteLine($"Field Type: {field.Type}");
            Console.WriteLine($"Field Value: {field.Result}");
            Console.WriteLine();
        }

        // (Optional) Modify a specific form field by name.
        FormField targetField = doc.Range.FormFields["CustomerName"];
        if (targetField != null)
        {
            targetField.Result = "John Doe";
        }

        // Save the modified document to a new file.
        string outputPath = @"C:\Docs\SampleForm_Modified.docx";
        doc.Save(outputPath);
    }
}
