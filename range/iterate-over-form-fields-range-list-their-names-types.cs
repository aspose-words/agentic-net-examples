using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

// Alias to avoid conflict with System.Range introduced in C# 8.
using AsposeRange = Aspose.Words.Range;

class ListFormFieldsInRange
{
    static void Main()
    {
        // Load an existing Word document.
        // Replace "input.docx" with the path to your document.
        Document doc = new Document("input.docx");

        // Define the range you want to inspect.
        // Here we use the whole document range, but you can use any sub‑range.
        AsposeRange range = doc.Range;

        // Get the collection of form fields that are present in the range.
        FormFieldCollection formFields = range.FormFields;

        // Iterate over the collection using its enumerator.
        using (IEnumerator<FormField> enumerator = formFields.GetEnumerator())
        {
            while (enumerator.MoveNext())
            {
                FormField field = enumerator.Current;
                // Output the field's name and its type (as a FieldType enum value).
                Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
            }
        }

        // Optionally, save the document after any modifications.
        // doc.Save("output.docx");
    }
}
