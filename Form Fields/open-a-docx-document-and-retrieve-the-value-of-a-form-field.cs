using System;
using Aspose.Words;
using Aspose.Words.Fields;

class RetrieveFormFieldValue
{
    static void Main()
    {
        // Load the existing DOCX document.
        // The Document constructor handles opening the file.
        Document doc = new Document("input.docx");

        // Access the collection of form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Example 1: Retrieve a form field by its bookmark (field) name.
        // Replace "MyTextInput" with the actual name of the form field you want.
        FormField fieldByName = formFields["MyTextInput"];
        if (fieldByName != null)
        {
            Console.WriteLine($"Field '{fieldByName.Name}' value: {fieldByName.Result}");
        }
        else
        {
            Console.WriteLine("Form field with the specified name was not found.");
        }

        // Example 2: Retrieve a form field by index (e.g., the first field).
        // Index is zero‑based; -1 returns the last field.
        FormField fieldByIndex = formFields[0];
        if (fieldByIndex != null)
        {
            Console.WriteLine($"Field at index 0 ('{fieldByIndex.Name}') value: {fieldByIndex.Result}");
        }
        else
        {
            Console.WriteLine("No form field found at the specified index.");
        }
    }
}
