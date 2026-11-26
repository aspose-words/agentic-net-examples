using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field with initial items.
        // The selectedIndex must be within the bounds of the items array.
        string[] comboItems = { "Option1", "Option2" };
        FormField combo = builder.InsertComboBox("DropDown1", comboItems, 0);
        builder.Writeln();

        // Ensure at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Save the document (required by the workflow).
        doc.Save("FormFields.docx");

        // Extract automatically generated bookmark names for all form fields.
        // In Aspose.Words, each legacy form field has a bookmark with the same name as the field.
        Dictionary<string, FormField> bookmarkLookup = new Dictionary<string, FormField>();

        foreach (FormField field in doc.Range.FormFields)
        {
            // Validate that the field has a name (bookmark name).
            if (!string.IsNullOrEmpty(field.Name))
            {
                // Store the mapping from bookmark name to the form field.
                bookmarkLookup[field.Name] = field;
            }
        }

        // Example usage: print the collected bookmark names.
        foreach (var kvp in bookmarkLookup)
        {
            Console.WriteLine($"Bookmark Name: {kvp.Key}, Field Type: {kvp.Value.Type}");
        }
    }
}
