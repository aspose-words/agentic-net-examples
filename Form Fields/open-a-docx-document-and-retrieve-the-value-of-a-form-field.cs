using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Path to the DOCX file that contains form fields.
        string filePath = "input.docx";

        // Load the document from the file system.
        Document doc = new Document(filePath);

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Example: retrieve the first form field by index.
        // You can also retrieve by name: doc.Range.FormFields["MyFormFieldName"]
        FormField formField = formFields[0];

        // The Result property holds the current value of the form field.
        string fieldValue = formField.Result;

        // Output the field name and its value.
        Console.WriteLine($"Form field '{formField.Name}' value: {fieldValue}");
    }
}
