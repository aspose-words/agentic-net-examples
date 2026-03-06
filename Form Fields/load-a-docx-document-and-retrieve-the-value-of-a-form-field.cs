using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Retrieve a form field.
        // You can access by index (zero‑based) or by bookmark name.
        // Example using index:
        FormField formField = doc.Range.FormFields[0];
        // Example using name (uncomment if you know the field name):
        // FormField formField = doc.Range.FormFields["MyFormFieldName"];

        // Get the current value of the form field.
        string fieldValue = formField.Result;

        // Output the value.
        Console.WriteLine($"Form field value: {fieldValue}");
    }
}
