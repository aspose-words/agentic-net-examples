using System;
using System.Collections.Generic;
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
        FormField textField = builder.InsertTextInput(
            name: "TextField1",
            type: TextFormFieldType.Regular,
            format: "",
            fieldValue: "Placeholder",
            maxLength: 0);
        // Set an explicit value for the text field.
        textField.SetTextInputValue("Hello World");

        // Insert a checkbox form field with a unique name.
        FormField checkBox = builder.InsertCheckBox(
            name: "CheckBox1",
            checkedValue: false,
            size: 0);

        // Insert a combo box (dropdown) form field with a unique name.
        FormField comboBox = builder.InsertComboBox(
            name: "ComboBox1",
            items: new[] { "Option A", "Option B", "Option C" },
            selectedIndex: 0);

        // Validate that each form field has a distinct name.
        FormFieldCollection formFields = doc.Range.FormFields;
        HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (FormField field in formFields)
        {
            if (!names.Add(field.Name))
            {
                throw new InvalidOperationException($"Duplicate form field name detected: {field.Name}");
            }
        }

        // Output the names of the created form fields.
        Console.WriteLine("Created form fields with unique names:");
        foreach (FormField field in formFields)
        {
            Console.WriteLine($"- {field.Name} (Type: {field.Type})");
        }

        // Save the document to the local file system.
        string outputPath = "FormFieldsUniqueNames.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
