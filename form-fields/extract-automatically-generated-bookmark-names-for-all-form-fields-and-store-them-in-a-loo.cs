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

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptTerms", false, 50);
        checkBox.CalculateOnExit = true;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field.
        builder.Write("Select a fruit: ");
        string[] fruits = { "Apple", "Banana", "Cherry" };
        FormField comboBox = builder.InsertComboBox("FruitChoice", fruits, 0);
        comboBox.CalculateOnExit = true;

        // Save the document containing the form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Extract automatically generated bookmark names for all form fields.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Dictionary to hold bookmark name -> form field mapping.
        var bookmarkLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);

        foreach (FormField field in formFields)
        {
            // Each form field automatically creates a bookmark with the same name.
            string bookmarkName = field.Name;
            if (!string.IsNullOrEmpty(bookmarkName))
                bookmarkLookup[bookmarkName] = field;
        }

        // Output the collected bookmark names.
        Console.WriteLine("Extracted bookmark names for form fields:");
        foreach (var kvp in bookmarkLookup)
        {
            Console.WriteLine($"Bookmark: \"{kvp.Key}\", Field Type: {kvp.Value.Type}");
        }

        // No further modifications are required, but ensure the document is saved again.
        doc.Save(outputPath);
    }
}
