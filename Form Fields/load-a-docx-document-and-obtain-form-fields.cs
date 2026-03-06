using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system using the Document(string) constructor.
        Document doc = new Document("SampleForm.docx");

        // Retrieve the collection of all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Display the total number of form fields found.
        Console.WriteLine($"Form fields count: {formFields.Count}");

        // Iterate through each form field and output its name and type.
        for (int i = 0; i < formFields.Count; i++)
        {
            FormField field = formFields[i];
            Console.WriteLine($"[{i}] Name: {field.Name}, Type: {field.Type}");
        }
    }
}
