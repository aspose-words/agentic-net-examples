using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class ExtractFormFieldBookmarks
{
    static void Main()
    {
        // Create a new document and add a form field programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert a text input form field. Aspose.Words will automatically create a bookmark for it.
        builder.InsertTextInput("MyField", TextFormFieldType.Regular, "", "Default value", 0);

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

        // Save the document (optional, just to demonstrate that the document is valid).
        doc.Save("OutputDocument.docx");
    }
}
