using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document from the file system.
        // This uses the Document(string) constructor, which is the provided load rule.
        Document doc = new Document("input.docx");

        // Retrieve the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Iterate through all form fields and output their names and current values.
        foreach (FormField field in formFields)
        {
            // The Result property holds the current value of the form field.
            // For a text input it is the entered text,
            // for a checkbox it is "True"/"False",
            // for a combo box it is the selected item.
            string value = field.Result;

            Console.WriteLine($"Form field name: {field.Name}");
            Console.WriteLine($"Form field type: {field.Type}");
            Console.WriteLine($"Form field value: {value}");
            Console.WriteLine();
        }

        // Example of accessing a specific form field by its bookmark/name.
        // Replace "MyTextInput" with the actual name of the field you want to read.
        FormField specificField = doc.Range.FormFields["MyTextInput"];
        if (specificField != null)
        {
            Console.WriteLine($"Specific field '{specificField.Name}' value: {specificField.Result}");
        }
    }
}
