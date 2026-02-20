using System;
using Aspose.Words;
using Aspose.Words.Fields;

class RetrieveFormFieldValue
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Specify the name of the form field whose value we want to retrieve.
        string fieldName = "MyTextInput";

        // Find the form field by name.
        FormField field = formFields[fieldName];

        if (field != null)
        {
            // For text input fields, the entered value is stored in the Result property.
            string fieldValue = field.Result;
            Console.WriteLine($"Form field \"{fieldName}\" value: {fieldValue}");
        }
        else
        {
            Console.WriteLine($"Form field \"{fieldName}\" not found.");
        }
    }
}
