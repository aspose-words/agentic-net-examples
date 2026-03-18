using System;
using Aspose.Words;
using Aspose.Words.Fields;
// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Words.Range;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Use the whole document as the target range.
        // If you need a specific sub‑range, obtain it via nodes (e.g., a paragraph.Range).
        AsposeRange range = doc.Range;

        // Locate the text input form field by its name.
        // The FormFields collection is accessible from the Range.
        FormField textField = range.FormFields["MyTextInput"];

        if (textField != null)
        {
            // Update the field's value.
            // SetTextInputValue applies the TextInputFormat (if any) and stores the result.
            textField.SetTextInputValue("New value");
        }
        else
        {
            Console.WriteLine("Form field 'MyTextInput' not found.");
        }

        // Recalculate any fields that depend on the form field's value.
        range.UpdateFields();

        // Save the modified document.
        doc.Save("output.docx");
    }
}
