using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractFormFieldBookmarks
{
    static void Main()
    {
        // Load an existing Word document that contains form fields.
        Document doc = new Document("InputDocument.docx");

        // Get the collection of all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Dictionary to hold bookmark name (key) and the corresponding form field (value).
        Dictionary<string, FormField> bookmarkLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);

        // Iterate through each form field and store its automatically generated bookmark name.
        foreach (FormField field in formFields)
        {
            // The Name property of a FormField is the bookmark name created by Aspose.Words.
            string bookmarkName = field.Name;

            if (!string.IsNullOrEmpty(bookmarkName))
            {
                // If duplicate names exist, the later one will overwrite the earlier entry.
                bookmarkLookup[bookmarkName] = field;
            }
        }

        // Example usage: print all extracted bookmark names.
        Console.WriteLine("Extracted bookmark names for form fields:");
        foreach (var kvp in bookmarkLookup)
        {
            Console.WriteLine($"Bookmark: {kvp.Key}, Field Type: {kvp.Value.Type}");
        }

        // Optionally, save the document after any modifications (none in this example).
        doc.Save("OutputDocument.docx");
    }
}
