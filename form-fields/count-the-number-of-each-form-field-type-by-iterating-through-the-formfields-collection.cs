using System;
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
        builder.Writeln();

        // Insert a checkbox form field.
        builder.InsertCheckBox("CheckBox1", false, 0);
        builder.Writeln();

        // Insert a dropdown (combo box) form field.
        string[] items = { "Option1", "Option2", "Option3" };
        builder.InsertComboBox("DropDown1", items, 0);
        builder.Writeln();

        // Save the document with the created form fields.
        string outputPath = "FormFields.docx";
        doc.Save(outputPath);

        // Ensure that at least one form field exists.
        if (doc.Range.FormFields.Count == 0)
        {
            throw new InvalidOperationException("The document does not contain any form fields.");
        }

        // Count each type of form field.
        int textInputCount = 0;
        int checkBoxCount = 0;
        int dropDownCount = 0;

        foreach (FormField field in doc.Range.FormFields)
        {
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
                    // Other field types are ignored for this example.
                    break;
            }
        }

        // Output the counts.
        Console.WriteLine($"Text Input Fields: {textInputCount}");
        Console.WriteLine($"Check Box Fields: {checkBoxCount}");
        Console.WriteLine($"Drop Down Fields: {dropDownCount}");
    }
}
