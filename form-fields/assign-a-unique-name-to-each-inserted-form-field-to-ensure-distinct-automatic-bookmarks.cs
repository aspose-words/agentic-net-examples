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

        // Counter to generate unique names for each form field.
        int fieldIndex = 1;

        // Insert a text input form field with a unique name.
        string textFieldName = $"TextField_{fieldIndex++}";
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(textFieldName, TextFormFieldType.Regular, "", "John Doe", 50);
        // Ensure the name is set (InsertTextInput already sets it, but we assign explicitly for clarity).
        textField.Name = textFieldName;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field with a unique name.
        string checkBoxName = $"CheckBox_{fieldIndex++}";
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(checkBoxName, false, 0);
        checkBox.Name = checkBoxName;

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field with a unique name.
        string comboBoxName = $"ComboBox_{fieldIndex++}";
        builder.Write("Select a country: ");
        string[] items = { "USA", "Canada", "Mexico" };
        FormField comboBox = builder.InsertComboBox(comboBoxName, items, 0);
        comboBox.Name = comboBoxName;

        // Verify that each form field has a distinct name.
        FormFieldCollection fields = doc.Range.FormFields;
        Console.WriteLine("Form fields in the document:");
        foreach (FormField field in fields)
        {
            Console.WriteLine($"- Name: {field.Name}, Type: {field.Type}");
        }

        // Save the document to disk.
        string outputPath = "FormFieldsUniqueNames.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
