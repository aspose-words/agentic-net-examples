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

        // Insert a text input form field with an initial unique name.
        FormField textField = builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Enter text", 0);
        // Insert a checkbox form field with an initial unique name.
        FormField checkBox = builder.InsertCheckBox("CheckBox", false, 0);
        // Insert a combo box (dropdown) form field with an initial unique name.
        string[] items = { "Option A", "Option B", "Option C" };
        FormField comboBox = builder.InsertComboBox("ComboBox", items, 0);

        // Ensure each form field has a distinct name (bookmark) by appending an index if needed.
        FormFieldCollection fields = doc.Range.FormFields;
        for (int i = 0; i < fields.Count; i++)
        {
            FormField field = fields[i];
            if (field == null) continue; // Safety check.

            // Build a unique name based on the field type and its position.
            string baseName = field.Type switch
            {
                FieldType.FieldFormTextInput => "TextField",
                FieldType.FieldFormCheckBox => "CheckBox",
                FieldType.FieldFormDropDown => "ComboBox",
                _ => "FormField"
            };

            string uniqueName = $"{baseName}_{i + 1}";
            field.Name = uniqueName; // Assign the unique name (also creates a bookmark with the same name).
        }

        // Optional: Output the assigned names to the console for verification.
        Console.WriteLine("Assigned form field names:");
        foreach (FormField field in doc.Range.FormFields)
        {
            Console.WriteLine($"{field.Type}: {field.Name}");
        }

        // Save the document to disk.
        string outputPath = "FormFieldsUnique.docx";
        doc.Save(outputPath);
    }
}
