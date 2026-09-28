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

        // Insert a text input form field with a unique name.
        builder.Writeln("Please enter your name:");
        FormField textField = builder.InsertTextInput(
            "TextField1",                     // Unique name
            TextFormFieldType.Regular,        // Field type
            "",                               // Format
            "John Doe",                       // Default text
            0);                               // Max length (0 = unlimited)

        // Insert a checkbox form field with a unique name.
        builder.Writeln();
        builder.Writeln("Accept terms:");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox1",                      // Unique name
            false,                            // Default state
            0);                               // Size (0 = default)

        // Insert a dropdown (combo box) form field with a unique name and items.
        builder.Writeln();
        builder.Writeln("Select country:");
        string[] countries = { "USA", "Canada", "Mexico" };
        FormField comboBox = builder.InsertComboBox(
            "DropDown1",                      // Unique name
            countries,                        // Items
            0);                               // Selected index (first item)

        // Ensure at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were created.");
        }

        // Validate that each form field has a distinct, non‑empty name.
        HashSet<string> fieldNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (FormField field in doc.Range.FormFields)
        {
            if (string.IsNullOrEmpty(field.Name))
                throw new InvalidOperationException("A form field has an empty name.");

            if (!fieldNames.Add(field.Name))
                throw new InvalidOperationException($"Duplicate form field name detected: {field.Name}");
        }

        // Save the document.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Output the names of the created form fields.
        Console.WriteLine("Form fields created with unique names:");
        foreach (FormField field in doc.Range.FormFields)
        {
            Console.WriteLine($"- {field.Name}");
        }
    }
}
