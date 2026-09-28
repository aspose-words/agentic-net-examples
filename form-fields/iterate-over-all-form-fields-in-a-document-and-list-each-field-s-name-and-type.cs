using System;
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
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln(); // Add a line break.

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBoxField", true, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field.
        string[] items = { "Option 1", "Option 2", "Option 3" };
        builder.InsertComboBox("DropDownField", items, 0);
        builder.Writeln();

        // Save the document with the created form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Ensure that the document contains at least one form field.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Iterate over all form fields and output their name and type.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Guard against null (should not happen, but follows nullable safety rules).
            if (field == null)
                continue;

            string fieldName = field.Name;
            string fieldType = field.Type.ToString();

            Console.WriteLine($"Field Name: {fieldName}, Field Type: {fieldType}");
        }
    }
}
