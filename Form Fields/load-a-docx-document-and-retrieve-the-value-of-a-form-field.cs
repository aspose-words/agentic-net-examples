using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load the DOCX document from file.
        Document doc = new Document("Input.docx");

        // Access the form field by its name (replace with your field's name).
        FormField formField = doc.Range.FormFields["MyFormField"];

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
