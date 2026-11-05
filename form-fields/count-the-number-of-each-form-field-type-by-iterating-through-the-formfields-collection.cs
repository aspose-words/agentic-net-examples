using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        // The InsertTextInput method requires a maxLength argument; use 50 as an example.
        builder.InsertTextInput("TextField", TextFormFieldType.Regular, "", "Default text", 50);

        // Insert a checkbox form field.
        builder.Writeln(); // Move to a new line.
        builder.InsertCheckBox("CheckBoxField", false, 0);

        // Insert a dropdown (combo box) form field.
        builder.Writeln(); // Move to a new line.
        builder.InsertComboBox("DropDownField", new string[] { "Option1", "Option2", "Option3" }, 0);

        // Save the document after creation.
        const string outputPath = "FormFieldsCount.docx";
        doc.Save(outputPath);

        // Ensure the document contains at least one form field.
        if (doc.Range?.FormFields == null || doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Counters for each form field type.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        // Iterate through the FormFields collection and count each type.
        foreach (FormField field in doc.Range.FormFields)
        {
            if (field == null) continue; // Guard against null entries.

            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    textInputCount++;
                    break;
                case FieldType.FieldFormCheckBox:
                    checkBoxCount++;
                    break;
                case FieldType.FieldFormDropDown:
                    dropDownCount++;
                    break;
                default:
                    // Other form field types are not expected in this example.
                    break;
            }
        }

        // Output the counts to the console.
        Console.WriteLine($"Text Input fields: {textInputCount}");
        Console.WriteLine($"Check Box fields: {checkBoxCount}");
        Console.WriteLine($"Drop Down fields: {dropDownCount}");
    }
}
