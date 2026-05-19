using System;
using System.Collections.Generic;
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

        // Insert a few different types of form fields, each with a unique name.
        builder.Write("Text input field: ");
        FormField textField = builder.InsertTextInput("TextField_1", TextFormFieldType.Regular, "", "Enter text here", 0);
        builder.InsertParagraph();

        builder.Write("Check box field: ");
        FormField checkBox = builder.InsertCheckBox("CheckBox_1", false, 0);
        builder.InsertParagraph();

        builder.Write("Combo box field: ");
        FormField comboBox = builder.InsertComboBox("ComboBox_1", new[] { "Option A", "Option B", "Option C" }, 0);
        builder.InsertParagraph();

        // Insert additional text input fields in a loop, ensuring each gets a distinct name.
        for (int i = 2; i <= 5; i++)
        {
            builder.Write($"Additional text field {i}: ");
            builder.InsertTextInput($"TextField_{i}", TextFormFieldType.Regular, "", $"Placeholder {i}", 0);
            builder.InsertParagraph();
        }

        // Validate that all form fields have unique names.
        FormFieldCollection formFields = doc.Range.FormFields;
        HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (FormField field in formFields)
        {
            if (field == null)
                continue; // Safety check.

            string name = field.Name;
            if (string.IsNullOrEmpty(name))
                throw new InvalidOperationException("A form field was found without a name.");

            if (!names.Add(name))
                throw new InvalidOperationException($"Duplicate form field name detected: {name}");
        }

        // Output the names of all form fields to the console.
        Console.WriteLine("Form fields and their unique names:");
        foreach (FormField field in formFields)
        {
            Console.WriteLine($"- {field.Type}: {field.Name}");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFields_UniqueNames.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
