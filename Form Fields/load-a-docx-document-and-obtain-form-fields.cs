using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from disk.
        Document doc = new Document("input.docx");

        // Retrieve the collection that contains all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Iterate through each form field and output basic information.
        foreach (FormField field in formFields)
        {
            Console.WriteLine($"Name: {field.Name}");
            Console.WriteLine($"Type: {field.Type}");
            Console.WriteLine($"Result: {field.Result}");
            Console.WriteLine(new string('-', 30));
        }
    }
}
