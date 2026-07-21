using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class FormFieldsDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a text input form field.
        builder.Write("Enter your name: ");
        FormField textField = builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "John Doe", 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a checkbox form field.
        builder.Write("Accept terms: ");
        FormField checkBoxField = builder.InsertCheckBox("AcceptTerms", false, 50);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Insert a combo box (drop‑down) form field.
        builder.Write("Select a fruit: ");
        string[] fruits = { "Apple", "Banana", "Cherry" };
        FormField comboBoxField = builder.InsertComboBox("FruitChoice", fruits, 0);
        builder.InsertBreak(BreakType.ParagraphBreak);

        // Access the collection of form fields.
        FormFieldCollection formFields = doc.Range.FormFields;
        if (formFields == null || formFields.Count == 0)
            throw new InvalidOperationException("No form fields were created.");

        // Iterate through each form field and handle it based on its type.
        foreach (FormField field in formFields)
        {
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    // Update the text input value.
                    field.SetTextInputValue("Alice Smith");
                    // Validate the update.
                    if (field.Result != "Alice Smith")
                        throw new InvalidOperationException("Failed to set text input value.");
                    break;

                case FieldType.FieldFormCheckBox:
                    // Check the box.
                    field.Checked = true;
                    // Validate the update.
                    if (!field.Checked)
                        throw new InvalidOperationException("Failed to check the checkbox.");
                    break;

                case FieldType.FieldFormDropDown:
                    // Select the second item ("Banana").
                    field.DropDownSelectedIndex = 1;
                    // Validate the update.
                    if (field.DropDownSelectedIndex != 1)
                        throw new InvalidOperationException("Failed to select dropdown item.");
                    break;

                default:
                    // Unexpected field type.
                    throw new NotSupportedException($"Unsupported form field type: {field.Type}");
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FormFieldsDemo.docx");
        doc.Save(outputPath);

        // Output a summary of the modified fields.
        Console.WriteLine("Form fields have been updated and saved to:");
        Console.WriteLine(outputPath);
        Console.WriteLine();
        Console.WriteLine("Field details after modification:");
        foreach (FormField field in formFields)
        {
            Console.Write($"- {field.Name} ({field.Type}): ");
            switch (field.Type)
            {
                case FieldType.FieldFormTextInput:
                    Console.WriteLine($"Result = \"{field.Result}\"");
                    break;
                case FieldType.FieldFormCheckBox:
                    Console.WriteLine($"Checked = {field.Checked}");
                    break;
                case FieldType.FieldFormDropDown:
                    Console.WriteLine($"Selected = \"{field.Result}\" (Index {field.DropDownSelectedIndex})");
                    break;
            }
        }
    }
}
