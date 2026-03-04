using System;
using Aspose.Words;
using Aspose.Words.Fields; // Added for FormField type

class LoadDocxForFormFields
{
    static void Main()
    {
        // Path to the DOCX file that contains form fields.
        string docPath = @"C:\Docs\SampleForm.docx";

        // Load the existing document from the file system.
        // This uses the Document(string) constructor, which is the prescribed load rule.
        Document doc = new Document(docPath);

        // Iterate over all form fields in the document.
        // The FormFields collection is accessed via the document's Range.
        foreach (FormField field in doc.Range.FormFields)
        {
            Console.WriteLine($"Field Name: {field.Name}, Type: {field.Type}");
            // You can read or modify the field value here, e.g.:
            // field.Result = "New value";
        }

        // (Optional) Save the document after modifications.
        // doc.Save(@"C:\Docs\SampleForm_Updated.docx");
    }
}
