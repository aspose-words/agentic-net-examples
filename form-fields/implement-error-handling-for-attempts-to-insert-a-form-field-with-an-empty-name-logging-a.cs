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

        // Insert a valid text input form field.
        InsertTextInput(builder, "UserName", TextFormFieldType.Regular, "", "Enter name", 50);

        // Attempt to insert a text input with an empty name – should log a warning.
        InsertTextInput(builder, "", TextFormFieldType.Regular, "", "Should be skipped", 30);

        // Insert a valid checkbox form field.
        InsertCheckBox(builder, "AcceptTerms", true, 20);

        // Attempt to insert a checkbox with an empty name – should log a warning.
        InsertCheckBox(builder, "", false, 20);

        // Insert a valid combo box (dropdown) form field.
        InsertComboBox(builder, "Country", new[] { "USA", "Canada", "Mexico" }, 0);

        // Attempt to insert a combo box with an empty name – should log a warning.
        InsertComboBox(builder, "", new[] { "Option1", "Option2" }, 0);

        // Save the document to the local file system.
        doc.Save("FormFields_Output.docx");
    }

    // Helper method for inserting a text input form field with validation.
    private static void InsertTextInput(DocumentBuilder builder, string name, TextFormFieldType type,
        string format, string fieldValue, int maxLength)
    {
        if (string.IsNullOrEmpty(name))
        {
            Console.WriteLine("Warning: Attempted to insert a text input form field with an empty name. Skipping insertion.");
            return;
        }

        builder.InsertTextInput(name, type, format, fieldValue, maxLength);
    }

    // Helper method for inserting a checkbox form field with validation.
    private static void InsertCheckBox(DocumentBuilder builder, string name, bool checkedValue, int size)
    {
        if (string.IsNullOrEmpty(name))
        {
            Console.WriteLine("Warning: Attempted to insert a checkbox form field with an empty name. Skipping insertion.");
            return;
        }

        builder.InsertCheckBox(name, checkedValue, size);
    }

    // Helper method for inserting a combo box (dropdown) form field with validation.
    private static void InsertComboBox(DocumentBuilder builder, string name, string[] items, int selectedIndex)
    {
        if (string.IsNullOrEmpty(name))
        {
            Console.WriteLine("Warning: Attempted to insert a combo box form field with an empty name. Skipping insertion.");
            return;
        }

        builder.InsertComboBox(name, items, selectedIndex);
    }
}
