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

        // Insert a text input form field with a unique name.
        FormField textField = builder.InsertTextInput("TextField1", TextFormFieldType.Regular, "", "Enter text here", 0);
        // Insert a checkbox form field with a unique name.
        FormField checkBox = builder.InsertCheckBox("CheckBox1", false, 0);
        // Insert a combo box (dropdown) form field with a unique name.
        FormField comboBox = builder.InsertComboBox("ComboBox1", new[] { "Option A", "Option B", "Option C" }, 0);

        // Insert additional text fields in a loop to guarantee unique names.
        for (int i = 2; i <= 5; i++)
        {
            builder.Writeln(); // Move to a new line.
            string fieldName = $"TextField{i}";
            builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", $"Placeholder {i}", 0);
        }

        // Ensure that at least one form field exists.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields.Count == 0)
        {
            throw new InvalidOperationException("No form fields were created.");
        }

        // Validate that each form field has a distinct name.
        var nameSet = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (FormField field in formFields)
        {
            if (string.IsNullOrEmpty(field.Name))
                throw new InvalidOperationException("A form field has an empty name.");

            if (!nameSet.Add(field.Name))
                throw new InvalidOperationException($"Duplicate form field name detected: {field.Name}");
        }

        // Output the names of the created form fields.
        Console.WriteLine("Created form fields with unique names:");
        foreach (FormField field in formFields)
        {
            Console.WriteLine($"- {field.Name} (Type: {field.Type})");
        }

        // Update fields (if any calculations are needed) and save the document.
        doc.UpdateFields();
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "FormFieldsUniqueNames.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
