using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        textField.CalculateOnExit = true;

        // Insert a checkbox form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 0);
        checkBox.CalculateOnExit = true;

        // Insert a combo box (dropdown) form field.
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Select a country: ");
        string[] countries = { "USA", "Canada", "Mexico" };
        FormField comboBox = builder.InsertComboBox("CountrySelect", countries, 0);
        comboBox.CalculateOnExit = true;

        // Save the document containing the form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Ensure that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Extract automatically generated bookmark names and store them in a lookup dictionary.
        // The bookmark name is the same as the form field's Name property.
        Dictionary<string, FormField> bookmarkLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);
        foreach (FormField field in formFields)
        {
            if (field == null)
                continue; // Guard against null entries.

            string bookmarkName = field.Name; // Automatically created bookmark name.
            if (!string.IsNullOrEmpty(bookmarkName) && !bookmarkLookup.ContainsKey(bookmarkName))
                bookmarkLookup.Add(bookmarkName, field);
        }

        // Output the collected bookmark names to the console.
        Console.WriteLine("Extracted bookmark names for form fields:");
        foreach (KeyValuePair<string, FormField> entry in bookmarkLookup)
        {
            Console.WriteLine($"- Bookmark: \"{entry.Key}\", Field Type: {entry.Value.Type}");
        }
    }
}
