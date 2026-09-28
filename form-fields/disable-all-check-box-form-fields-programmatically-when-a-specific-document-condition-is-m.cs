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

        // Insert three checkbox form fields with distinct names.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBox2", true, 0);
        builder.Writeln();
        builder.InsertCheckBox("CheckBox3", false, 0);
        builder.Writeln();

        // Save the initial document (optional, shows the state before disabling).
        doc.Save("initial.docx");

        // Condition that determines whether checkboxes should be disabled.
        bool disableCheckBoxes = true; // Change as needed.

        if (disableCheckBoxes)
        {
            // Iterate through all form fields in the document.
            foreach (FormField field in doc.Range.FormFields)
            {
                // Only process checkbox fields.
                if (field.Type == FieldType.FieldFormCheckBox)
                {
                    // Disable the checkbox so it cannot be edited in the UI.
                    field.Enabled = false;
                }
            }
        }

        // Save the modified document where checkboxes are disabled.
        doc.Save("output.docx");
    }
}
