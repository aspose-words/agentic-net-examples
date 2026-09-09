using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Counter to generate unique names.
        int fieldIndex = 1;

        // Insert a text input form field with a unique name.
        string textFieldName = $"TextField_{fieldIndex++}";
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput(textFieldName, TextFormFieldType.Regular, "", "John Doe", 50);
        // Ensure the name was set.
        if (string.IsNullOrEmpty(textField.Name))
            throw new InvalidOperationException("Text input field name was not assigned.");

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field with a unique name.
        string checkBoxName = $"CheckBox_{fieldIndex++}";
        builder.Write("Accept terms: ");
        FormField checkBox = builder.InsertCheckBox(checkBoxName, false, 0);
        if (string.IsNullOrEmpty(checkBox.Name))
            throw new InvalidOperationException("Check box field name was not assigned.");

        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (dropdown) form field with a unique name.
        string comboBoxName = $"ComboBox_{fieldIndex++}";
        builder.Write("Select a country: ");
        string[] items = { "USA", "Canada", "Mexico" };
        FormField comboBox = builder.InsertComboBox(comboBoxName, items, 0);
        if (string.IsNullOrEmpty(comboBox.Name))
            throw new InvalidOperationException("Combo box field name was not assigned.");

        // Validate that all form fields exist in the collection.
        FormFieldCollection fields = doc.Range.FormFields;
        if (fields.Count != 3)
            throw new InvalidOperationException("Expected three form fields in the document.");

        // Output the names of the form fields to verify uniqueness.
        Console.WriteLine("Form fields and their unique names:");
        foreach (FormField field in fields)
        {
            Console.WriteLine($"- {field.Type}: {field.Name}");
        }

        // Save the document. Each form field automatically creates a bookmark with the same name.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFields_UniqueNames.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
