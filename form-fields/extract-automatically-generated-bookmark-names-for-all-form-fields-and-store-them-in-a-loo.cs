using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln(); // Add a line break.

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field.
        string[] items = { "Option1", "Option2", "Option3" };
        builder.InsertComboBox("DropDown1", items, 0);
        builder.Writeln();

        // Save the document with the created form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Extract automatically generated bookmark names for all form fields.
        var formFieldLookup = new Dictionary<string, FormField>(StringComparer.OrdinalIgnoreCase);

        foreach (FormField field in doc.Range.FormFields)
        {
            // Ensure the field has a valid name (bookmark name).
            if (!string.IsNullOrEmpty(field.Name))
            {
                // Add to the lookup dictionary.
                formFieldLookup[field.Name] = field;
            }
        }

        // Validate that at least one form field was found.
        if (formFieldLookup.Count == 0)
        {
            throw new InvalidOperationException("No form fields were found in the document.");
        }

        // Output the extracted bookmark names.
        Console.WriteLine("Extracted bookmark names for form fields:");
        foreach (KeyValuePair<string, FormField> kvp in formFieldLookup)
        {
            Console.WriteLine($"- {kvp.Key}");
        }

        // Demonstrate accessing a field via the lookup.
        if (formFieldLookup.TryGetValue("TextField1", out FormField textField))
        {
            Console.WriteLine($"Value of '{textField.Name}': {textField.Result}");
        }

        // Save the document again (optional, demonstrates that modifications could be persisted).
        doc.Save("FormFields_Processed.docx");
    }
}
