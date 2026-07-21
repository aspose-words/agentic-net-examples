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

        // Insert a combo box form field.
        builder.Write("Choose a value: ");
        FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "One", "Two", "Three" }, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a check box form field.
        builder.Write("Check this box: ");
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a text input form field.
        builder.Write("Enter text: ");
        FormField textInput = builder.InsertTextInput("MyTextInput", TextFormFieldType.Regular, "", "Placeholder", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Retrieve the collection of form fields in the whole document range.
        FormFieldCollection formFields = doc.Range.FormFields;

        // Get the count of form fields.
        int count = formFields.Count;

        // Output the count to the console.
        Console.WriteLine($"Form fields count: {count}");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document (optional, but fulfills the save requirement).
        string outputPath = Path.Combine(outputDir, "FormFieldsCount.docx");
        doc.Save(outputPath);
    }
}
