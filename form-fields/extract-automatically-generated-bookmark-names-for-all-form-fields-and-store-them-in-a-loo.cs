using System;
using System.Collections.Generic;
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
            "NameField",                     // field name (bookmark will be created with this name)
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
        string[] countries = { "USA", "Canada", "Mexico" };
        FormField comboBox = builder.InsertComboBox(
            "Country",       // field name
            countries,       // items
            0);              // selected index

        // Save the document so that the form fields (and their bookmarks) are persisted.
        doc.Save("FormFields.docx");

        // Ensure that at least one form field exists before processing.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Extract automatically generated bookmark names (they match the form field names)
        // and store them in a lookup dictionary.
        var bookmarkLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);
        foreach (FormField field in formFields)
        {
            // The Name property is the bookmark name created by the Insert* methods.
            if (!string.IsNullOrEmpty(field.Name))
                bookmarkLookup[field.Name] = field;
        }

        // Demonstrate the lookup dictionary by printing each bookmark name and its field type.
        foreach (KeyValuePair<string, FormField> entry in bookmarkLookup)
        {
            Console.WriteLine($"Bookmark: {entry.Key}, Field Type: {entry.Value.Type}");
        }
    }
}
