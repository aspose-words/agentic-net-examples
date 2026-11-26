using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Default text", 0);
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        // Insert a combo box (dropdown) form field with initial items.
        string[] comboItems = { "Option A", "Option B", "Option C" };
        builder.InsertComboBox("ComboBox1", comboItems, 0);
        builder.Writeln();

        // Save the document with the created form fields.
        const string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Iterate over all form fields and list each field's name and type.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Ensure the field is not null (defensive check).
            if (field != null)
            {
                Console.WriteLine($"Name: {field.Name}, Type: {field.Type}");
            }
        }

        // Save again to ensure the document is persisted after reading.
        doc.Save(outputPath);
    }
}
