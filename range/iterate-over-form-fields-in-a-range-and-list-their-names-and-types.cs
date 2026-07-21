using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a combo box form field.
        builder.Write("Choose a value: ");
        builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this box: ");
        builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text: ");
        builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);

        // Get the range that covers the whole document.
        // Use the fully qualified Aspose.Words.Range to avoid ambiguity with System.Range.
        Aspose.Words.Range range = doc.Range;

        // Iterate over all form fields in the range and output their name and type.
        foreach (FormField field in range.FormFields)
        {
            Console.WriteLine($"Form field name: {field.Name}, type: {field.Type}");
        }
    }
}
