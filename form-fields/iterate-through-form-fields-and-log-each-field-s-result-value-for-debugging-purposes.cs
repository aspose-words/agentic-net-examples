using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for inserting form fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        FormField textField = builder.InsertTextInput(
            "TextField1",                     // field name
            TextFormFieldType.Regular,        // field type
            "",                               // format (none)
            "Default text",                   // default text
            0);                               // max length (0 = unlimited)

        // Insert a checkbox form field.
        FormField checkBox = builder.InsertCheckBox(
            "CheckBox1",                      // field name
            true,                             // default state (checked)
            0);                               // size (default)

        // Insert a dropdown (combo box) form field with predefined items.
        // The overload requires the field name, an array of item texts, and the selected index.
        FormField comboBox = builder.InsertComboBox(
            "DropDown1",                       // field name
            new[] { "Option 1", "Option 2", "Option 3" }, // items
            0);                                // selected index (first item)

        // Save the document with the created form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Load the document (demonstrates reading from a file).
        Document loadedDoc = new Document(outputPath);

        // Ensure that at least one form field exists.
        if (loadedDoc.Range.FormFields == null || loadedDoc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were found in the document.");
        }

        // Iterate through all form fields and log their Result values.
        foreach (FormField field in loadedDoc.Range.FormFields)
        {
            // Guard against null (should not happen in the collection).
            if (field == null)
                continue;

            // For checkboxes, Result returns "True"/"False".
            // For dropdowns, Result returns the selected item text.
            string result = field.Result ?? string.Empty;

            Console.WriteLine($"Field Name: {field.Name}, Result: {result}");
        }

        // Optionally, save the document again after reading (no changes made here).
        loadedDoc.Save("FormFields_Processed.docx");
    }
}
