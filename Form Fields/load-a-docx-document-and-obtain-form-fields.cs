using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system using the Document(string) constructor.
        Document doc = new Document("input.docx");

        // Retrieve the collection of all form fields in the document's range.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Display the total number of form fields found.
        Console.WriteLine($"Form fields count: {formFields.Count}");

        // Iterate through each form field and output its name and type.
        foreach (FormField field in formFields)
        {
            Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
        }
    }
}
