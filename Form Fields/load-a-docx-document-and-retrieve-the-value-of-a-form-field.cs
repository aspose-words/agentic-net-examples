using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the DOCX file that contains the form fields.
        string docPath = "FormFields.docx";

        // Load the document from the file system using the Document(string) constructor.
        Document doc = new Document(docPath);

        // Retrieve a form field by its bookmark name (the field's Name property).
        // Replace "MyTextInput" with the actual name of the form field you need.
        FormField formField = doc.Range.FormFields["MyTextInput"];

        if (formField != null)
        {
            // The Result property holds the current value entered in the form field.
            string fieldValue = formField.Result;
            Console.WriteLine($"Form field '{formField.Name}' value: {fieldValue}");
        }
        else
        {
            Console.WriteLine("Form field not found.");
        }
    }
}
