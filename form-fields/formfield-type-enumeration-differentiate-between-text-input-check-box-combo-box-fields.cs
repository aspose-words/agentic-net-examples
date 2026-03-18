using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class FormFieldTypeDemo
{
    static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box (dropdown) form field.
        builder.Write("Select a fruit: ");
        FormField comboBox = builder.InsertComboBox("FruitCombo", new[] { "Apple", "Banana", "Cherry" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox("TermsCheck", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textInput = builder.InsertTextInput("NameInput", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Iterate through all form fields and differentiate by their Type.
        foreach (FormField field in doc.Range.FormFields)
        {
            // The Type property returns a FieldType enumeration value.
            switch (field.Type)
            {
                case FieldType.FieldFormDropDown:
                    Console.WriteLine($"ComboBox: Name=\"{field.Name}\", Selected=\"{field.Result}\"");
                    break;

                case FieldType.FieldFormCheckBox:
                    // For a check box, the Checked property indicates its state.
                    Console.WriteLine($"CheckBox: Name=\"{field.Name}\", Checked={field.Checked}");
                    break;

                case FieldType.FieldFormTextInput:
                    // For a text input, the Result property holds the entered text.
                    Console.WriteLine($"TextInput: Name=\"{field.Name}\", Text=\"{field.Result}\"");
                    break;

                default:
                    Console.WriteLine($"Other field type: {field.Type}");
                    break;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "FormFieldsDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
