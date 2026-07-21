using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and insert several form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text input form field.
        builder.Write("Enter your name: ");
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.Writeln();

        // Checkbox form field.
        builder.Write("Accept terms: ");
        builder.InsertCheckBox("AcceptCheck", false, 50);
        builder.Writeln();

        // Dropdown (combo box) form field.
        builder.Write("Select country: ");
        string[] countries = { "USA", "Canada", "Mexico" };
        builder.InsertComboBox("CountryCombo", countries, 0);
        builder.Writeln();

        // Save the document with the created form fields.
        const string filePath = "FormFields.docx";
        doc.Save(filePath);

        // Load the document and extract automatically generated bookmark names.
        Document loadedDoc = new Document(filePath);
        FormFieldCollection formFields = loadedDoc.Range.FormFields;

        // Dictionary to map bookmark name (same as form field name) to the form field.
        Dictionary<string, FormField> bookmarkLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);

        foreach (FormField field in formFields)
        {
            // The bookmark name is automatically created with the same name as the form field.
            if (!string.IsNullOrEmpty(field.Name))
            {
                bookmarkLookup[field.Name] = field;
            }
        }

        // Output the collected bookmark names and their field types.
        Console.WriteLine($"Found {bookmarkLookup.Count} form fields with automatically generated bookmarks:");
        foreach (KeyValuePair<string, FormField> entry in bookmarkLookup)
        {
            Console.WriteLine($"Bookmark: \"{entry.Key}\", Field Type: {entry.Value.Type}");
        }
    }
}
