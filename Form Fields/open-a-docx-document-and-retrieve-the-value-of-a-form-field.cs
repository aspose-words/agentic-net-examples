using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("Input.docx");

        // Get the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Retrieve a specific form field by its name (bookmark name).
        // Replace "MyTextInput" with the actual name of the form field you want to read.
        FormField field = formFields["MyTextInput"];

        if (field != null)
        {
            // The value entered by the user is stored in the Result property.
            string fieldValue = field.Result;
            Console.WriteLine($"Form field '{field.Name}' value: {fieldValue}");
        }
        else
        {
            Console.WriteLine("Form field not found.");
        }
    }
}
