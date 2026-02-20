using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Get a form field by its zero‑based index.
        // Ensure the index is within the collection bounds.
        if (formFields.Count > 0)
        {
            FormField firstField = formFields[0];
            Console.WriteLine($"First field name: {firstField.Name}");
        }

        // Get a form field by its name.
        // Replace "MyFieldName" with the actual name of the field you want.
        string fieldName = "MyFieldName";
        FormField namedField = formFields[fieldName];
        if (namedField != null)
        {
            Console.WriteLine($"Field \"{fieldName}\" found. Type: {namedField.Type}");
        }
        else
        {
            Console.WriteLine($"Field \"{fieldName}\" not found.");
        }

        // (Optional) Save the document after any modifications.
        // doc.Save("Output.docx");
    }
}
