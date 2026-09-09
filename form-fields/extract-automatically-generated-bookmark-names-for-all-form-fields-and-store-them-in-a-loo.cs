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
            "TextField1",                     // field name (bookmark will be created with this name)
            TextFormFieldType.Regular,        // type of text field
            "",                               // default text
            "John Doe",                       // placeholder text
            30);                              // maximum length

        // Insert a checkbox form field.
        builder.Writeln(); // move to next line
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox1",                      // field name
            false,                            // default unchecked
            50);                              // size in points

        // Insert a dropdown (combo box) form field.
        builder.Writeln();
        builder.Write("Select a fruit: ");
        FormField comboBox = builder.InsertComboBox(
            "DropDown1",                      // field name
            new[] { "Apple", "Banana", "Cherry" }, // items
            0);                               // default selected index

        // Save the document containing the form fields.
        const string docPath = "FormFields.docx";
        doc.Save(docPath);

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Ensure that at least one form field exists.
        if (formFields.Count == 0)
            throw new InvalidOperationException("The document does not contain any form fields.");

        // Dictionary to hold the mapping: form field name -> automatically generated bookmark name.
        // For legacy form fields the bookmark name is identical to the field name.
        Dictionary<string, string> fieldBookmarkLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (FormField field in formFields)
        {
            // Validate that the field has a name.
            if (string.IsNullOrEmpty(field.Name))
                continue; // Skip unnamed fields (should not happen with Insert* methods).

            // The bookmark created by Insert* methods uses the same name as the form field.
            string bookmarkName = field.Name;

            // Store the mapping.
            fieldBookmarkLookup[field.Name] = bookmarkName;
        }

        // Output the lookup dictionary to the console.
        Console.WriteLine("Form field to bookmark mapping:");
        foreach (KeyValuePair<string, string> kvp in fieldBookmarkLookup)
        {
            Console.WriteLine($"Field Name: \"{kvp.Key}\"  ->  Bookmark Name: \"{kvp.Value}\"");
        }
    }
}
