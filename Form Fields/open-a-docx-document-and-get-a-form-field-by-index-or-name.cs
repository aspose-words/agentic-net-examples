using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Get a form field by its zero‑based index.
        // Returns null if the index is out of range.
        FormField fieldByIndex = doc.Range.FormFields[0];
        if (fieldByIndex != null)
        {
            Console.WriteLine($"Field at index 0: Name = {fieldByIndex.Name}, Result = {fieldByIndex.Result}");
        }

        // Get a form field by its bookmark (field) name.
        // Returns null if a field with the specified name does not exist.
        FormField fieldByName = doc.Range.FormFields["MyComboBox"];
        if (fieldByName != null)
        {
            Console.WriteLine($"Field named 'MyComboBox': Type = {fieldByName.Type}, Result = {fieldByName.Result}");
        }

        // Optionally save the document (e.g., after modifications).
        doc.Save("OutputDocument.docx");
    }
}
