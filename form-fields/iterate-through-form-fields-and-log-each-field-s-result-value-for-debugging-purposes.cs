using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter name: ");
        FormField textField = builder.InsertTextInput(
            "NameField",                     // field name
            TextFormFieldType.Regular,       // field type
            "",                              // default text (empty)
            "John Doe",                      // placeholder text
            50);                             // maximum length

        // Insert a checkbox form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "AcceptTerms",   // field name
            false,           // default unchecked
            50);             // size in points

        // Insert a combo box (dropdown) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select country: ");
        FormField comboBox = builder.InsertComboBox(
            "CountryField",                     // field name
            new[] { "USA", "Canada", "Mexico" }, // items
            0);                                 // default selected index

        // Validate that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were created in the document.");
        }

        // Iterate through all form fields and log each field's Result value.
        foreach (FormField field in formFields)
        {
            // Guard against a possible null entry.
            if (field != null)
            {
                // For debugging purposes, output the field name, type, and its current result.
                Console.WriteLine($"Field Name: {field.Name}, Type: {field.Type}, Result: {field.Result}");
            }
        }

        // Save the document to disk.
        doc.Save("FormFields.docx");
    }
}
