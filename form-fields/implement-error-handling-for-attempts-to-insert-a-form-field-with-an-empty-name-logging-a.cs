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

        // Helper method that validates the field name before insertion.
        void InsertFormField(Action insertAction, string fieldName)
        {
            if (string.IsNullOrEmpty(fieldName))
            {
                // Log a warning and skip the insertion.
                Console.WriteLine("Warning: Attempted to insert a form field with an empty name. Skipping insertion.");
                return;
            }

            // Name is valid – perform the insertion.
            insertAction();
        }

        // Insert a text input form field with a proper name.
        InsertFormField(
            () => builder.InsertTextInput("ValidText", TextFormFieldType.Regular, "", "Enter text here", 50),
            "ValidText");

        // Attempt to insert a checkbox with an empty name – should trigger the warning.
        InsertFormField(
            () => builder.InsertCheckBox("", false, 20),
            "");

        // Attempt to insert a combo box with a null name – also triggers the warning.
        string nullName = null;
        InsertFormField(
            () => builder.InsertComboBox(nullName, new[] { "Option1", "Option2" }, 0),
            nullName);

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFieldsExample.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
