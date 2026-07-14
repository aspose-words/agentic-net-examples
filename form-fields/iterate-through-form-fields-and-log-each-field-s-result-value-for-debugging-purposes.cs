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
        builder.Write("Enter your name: ");
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
            "AcceptTerms",                   // field name
            false,                           // default unchecked
            15);                             // size in points

        // Insert a dropdown (combo box) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a country: ");
        FormField comboBox = builder.InsertComboBox(
            "CountryBox",                    // field name
            new[] { "USA", "Canada", "UK" }, // items
            0);                              // default selected index

        // Ensure that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate through all form fields and log their Result values.
        foreach (FormField field in formFields)
        {
            // Guard against null (should not happen, but follows nullable safety rules).
            if (field != null)
            {
                // For checkbox fields, Result is "1" (checked) or "0" (unchecked).
                // For text and dropdown fields, Result contains the displayed text.
                Console.WriteLine($"Field Name: {field.Name}, Result: {field.Result}");
            }
        }

        // Save the document (required by the validation rules).
        doc.Save("FormFields.docx");
    }
}
