using System;
using Aspose.Words;
using Aspose.Words.Fields;

class FormFieldExample
{
    static void Main()
    {
        // Load an existing DOCX document from the file system.
        Document doc = new Document("Input.docx");

        // Retrieve the collection that contains all form fields in the document.
        FormFieldCollection formFields = doc.Range.FormFields;

        // ----- Access a form field by index -----
        // The indexer is zero‑based; negative values count from the end.
        FormField firstField = formFields[0];
        if (firstField != null)
        {
            Console.WriteLine($"Field at index 0: Name = {firstField.Name}, Type = {firstField.Type}");
        }

        // ----- Access a form field by name (bookmark) -----
        // The name lookup is case‑insensitive. Replace "MyCheckBox" with an actual field name.
        FormField namedField = formFields["MyCheckBox"];
        if (namedField != null)
        {
            Console.WriteLine($"Field named 'MyCheckBox': Name = {namedField.Name}, Type = {namedField.Type}");
        }

        // Save the document if any changes were made (optional).
        doc.Save("Output.docx");
    }
}
