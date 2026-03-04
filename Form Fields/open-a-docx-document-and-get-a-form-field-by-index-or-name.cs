using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // ----- Retrieve a form field by index -----
        // The collection is zero‑based; negative indexes count from the end.
        FormField fieldByIndex = doc.Range.FormFields[0];
        if (fieldByIndex != null)
        {
            Console.WriteLine($"Field at index 0:");
            Console.WriteLine($"  Name : {fieldByIndex.Name}");
            Console.WriteLine($"  Type : {fieldByIndex.Type}");
        }

        // ----- Retrieve a form field by name (bookmark name) -----
        // The name lookup is case‑insensitive.
        FormField fieldByName = doc.Range.FormFields["MyCheckBox"];
        if (fieldByName != null)
        {
            Console.WriteLine($"Field with name 'MyCheckBox':");
            Console.WriteLine($"  Name : {fieldByName.Name}");
            Console.WriteLine($"  Type : {fieldByName.Type}");
        }

        // Save the document if any changes were made.
        doc.Save("Output.docx");
    }
}
