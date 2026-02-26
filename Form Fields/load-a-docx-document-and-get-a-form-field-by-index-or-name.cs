using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // ----- Access a form field by index -----
        // The indexer is zero‑based; negative values count from the end.
        FormField fieldByIndex = doc.Range.FormFields[0]; // first field
        if (fieldByIndex != null)
        {
            Console.WriteLine($"Field at index 0: Name = \"{fieldByIndex.Name}\", Type = {fieldByIndex.Type}");
        }

        // Example of using a negative index to get the last field.
        FormField lastField = doc.Range.FormFields[-1];
        if (lastField != null)
        {
            Console.WriteLine($"Last field (index -1): Name = \"{lastField.Name}\", Type = {lastField.Type}");
        }

        // ----- Access a form field by name -----
        // The name lookup is case‑insensitive.
        FormField fieldByName = doc.Range.FormFields["MyCheckBox"];
        if (fieldByName != null)
        {
            Console.WriteLine($"Field named \"MyCheckBox\": Index = (unknown), Type = {fieldByName.Type}");
        }
        else
        {
            Console.WriteLine("Form field \"MyCheckBox\" not found.");
        }

        // Save the document if any changes were made (optional).
        doc.Save("Output.docx");
    }
}
