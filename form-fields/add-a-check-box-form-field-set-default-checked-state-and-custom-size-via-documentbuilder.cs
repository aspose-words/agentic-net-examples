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

        // Insert a description before the checkbox.
        builder.Write("Custom sized checked checkbox: ");

        // Insert a checkbox form field.
        // Parameters: name, default value, current checked state, size in points.
        FormField checkBox = builder.InsertCheckBox("MyCheckBox", true, true, 30);

        // Enable exact size so the specified size is applied.
        checkBox.IsCheckBoxExactSize = true;

        // Save the document to disk.
        const string outputPath = "CheckBoxFormField.docx";
        doc.Save(outputPath);

        // Verify that the field exists and its properties are set as expected.
        FormField? retrieved = doc.Range.FormFields["MyCheckBox"];
        if (retrieved != null)
        {
            Console.WriteLine($"Field '{retrieved.Name}' found.");
            Console.WriteLine($"Checked: {retrieved.Checked}");
            Console.WriteLine($"Default: {retrieved.Default}");
            Console.WriteLine($"Size (points): {retrieved.CheckBoxSize}");
        }
        else
        {
            throw new InvalidOperationException("The expected form field was not found.");
        }
    }
}
