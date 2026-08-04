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
            "UserName",                     // field name
            TextFormFieldType.Regular,      // field type
            "",                             // format (none)
            "John Doe",                     // placeholder text
            50);                            // max length

        // Insert a checkbox form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "AcceptTerms",                  // field name
            false,                          // default unchecked
            15);                            // size in points

        // Insert a combo box (dropdown) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a country: ");
        FormField comboBox = builder.InsertComboBox(
            "Country",                      // field name
            new[] { "USA", "Canada", "UK" },// items
            0);                             // default selected index

        // Ensure that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Iterate through all form fields and log their Result values.
        foreach (FormField field in formFields)
        {
            // Guard against null Result.
            string result = field.Result ?? string.Empty;
            Console.WriteLine($"Field Name: {field.Name}, Result: \"{result}\"");
        }

        // Save the document (required by the feature rules).
        doc.Save("FormFieldsOutput.docx");
    }
}
