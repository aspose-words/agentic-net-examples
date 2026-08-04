using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Attempt to insert a text input form field with a valid name.
        InsertTextInputSafe(builder, "ValidTextField", TextFormFieldType.Regular, "", "Enter text", 50);

        // Attempt to insert a text input form field with an empty name – should log a warning.
        InsertTextInputSafe(builder, "", TextFormFieldType.Regular, "", "Should not be added", 50);

        // Attempt to insert a checkbox with a valid name.
        InsertCheckBoxSafe(builder, "ValidCheckBox", false, 20);

        // Attempt to insert a checkbox with an empty name – should log a warning.
        InsertCheckBoxSafe(builder, "", true, 20);

        // Attempt to insert a combo box with a valid name.
        InsertComboBoxSafe(builder, "ValidComboBox", new[] { "Option1", "Option2" }, 0);

        // Attempt to insert a combo box with an empty name – should log a warning.
        InsertComboBoxSafe(builder, "", new[] { "OptionA", "OptionB" }, 0);

        // Save the document.
        doc.Save("FormFieldsOutput.docx");
    }

    // Inserts a text input form field only if the name is not empty.
    private static void InsertTextInputSafe(DocumentBuilder builder, string name, TextFormFieldType type,
                                            string format, string fieldValue, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            Console.WriteLine("Warning: Attempted to insert a text input form field with an empty name. Insertion skipped.");
            return;
        }

        builder.InsertTextInput(name, type, format, fieldValue, maxLength);
    }

    // Inserts a checkbox form field only if the name is not empty.
    private static void InsertCheckBoxSafe(DocumentBuilder builder, string name, bool defaultValue, int size)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            Console.WriteLine("Warning: Attempted to insert a checkbox form field with an empty name. Insertion skipped.");
            return;
        }

        // Using the overload that sets default and current checked values.
        builder.InsertCheckBox(name, defaultValue, defaultValue, size);
    }

    // Inserts a combo box form field only if the name is not empty.
    private static void InsertComboBoxSafe(DocumentBuilder builder, string name, string[] items, int selectedIndex)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            Console.WriteLine("Warning: Attempted to insert a combo box form field with an empty name. Insertion skipped.");
            return;
        }

        builder.InsertComboBox(name, items, selectedIndex);
    }
}
