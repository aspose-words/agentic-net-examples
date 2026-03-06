using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file that contains form fields.
        Document doc = new Document("FormFields.docx");

        // ----- Retrieve a form field by its zero‑based index -----
        // The FormFieldCollection indexer returns null if the index is out of range.
        FormField fieldByIndex = doc.Range.FormFields[0];
        if (fieldByIndex != null)
        {
            Console.WriteLine($"Field at index 0:");
            Console.WriteLine($"  Name : {fieldByIndex.Name}");
            Console.WriteLine($"  Type : {fieldByIndex.Type}");
        }

        // ----- Retrieve a form field by its bookmark (field) name -----
        // The name is case‑insensitive; returns null if not found.
        FormField fieldByName = doc.Range.FormFields["MyCheckBox"];
        if (fieldByName != null)
        {
            Console.WriteLine($"Field named 'MyCheckBox':");
            Console.WriteLine($"  Type    : {fieldByName.Type}");
            Console.WriteLine($"  Checked : {fieldByName.Checked}");
        }

        // Example modification: change the result of a text input field.
        FormField textField = doc.Range.FormFields["MyTextInput"];
        if (textField != null)
        {
            textField.Result = "Updated value";
        }

        // Save the document after any modifications.
        doc.Save("FormFields_Updated.docx");
    }
}
