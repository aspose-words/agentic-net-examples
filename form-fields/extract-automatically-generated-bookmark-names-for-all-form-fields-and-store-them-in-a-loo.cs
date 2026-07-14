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

        // Insert a combo box form field.
        builder.Write("Choose a fruit: ");
        FormField comboBox = builder.InsertComboBox("FruitCombo", new[] { "Apple", "Banana", "Cherry" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("AcceptCheck", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textInput = builder.InsertTextInput("NameInput", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Save the document (optional, demonstrates lifecycle rule).
        doc.Save("FormFields.docx");

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Validate that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("No form fields were found in the document.");

        // Create a lookup dictionary where the key is the automatically generated bookmark name.
        Dictionary<string, FormField> bookmarkLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);

        // Populate the dictionary.
        foreach (FormField field in formFields)
        {
            // The name of the form field is also the name of the automatically created bookmark.
            string bookmarkName = field.Name;

            // Ensure the bookmark actually exists in the document.
            if (doc.Range.Bookmarks[bookmarkName] == null)
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' was not found.");

            bookmarkLookup[bookmarkName] = field;
        }

        // Example usage: print each bookmark name and its form field type.
        foreach (KeyValuePair<string, FormField> entry in bookmarkLookup)
        {
            Console.WriteLine($"Bookmark: {entry.Key}, FormField Type: {entry.Value.Type}");
        }

        // Save the document again after processing (demonstrates saving after modifications if any).
        doc.Save("FormFields_Processed.docx");
    }
}
