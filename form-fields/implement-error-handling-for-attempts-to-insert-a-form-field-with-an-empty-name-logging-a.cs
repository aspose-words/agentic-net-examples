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

        // Attempt to insert a text input form field with an empty name.
        InsertFormField(builder, "", () =>
            builder.InsertTextInput("", TextFormFieldType.Regular, "", "Placeholder for empty name", 0));

        // Insert a valid text input form field.
        InsertFormField(builder, "ValidTextField", () =>
            builder.InsertTextInput("ValidTextField", TextFormFieldType.Regular, "", "Default text", 50));

        // Attempt to insert a checkbox form field with an empty name.
        InsertFormField(builder, "", () =>
            builder.InsertCheckBox("", false, 20));

        // Insert a valid checkbox form field.
        InsertFormField(builder, "ValidCheckBox", () =>
            builder.InsertCheckBox("ValidCheckBox", true, 20));

        // Save the document.
        doc.Save("FormFieldsOutput.docx");
    }

    // Helper method that validates the form field name before insertion.
    private static void InsertFormField(DocumentBuilder builder, string name, Action insertAction)
    {
        if (string.IsNullOrEmpty(name))
        {
            // Log a warning and skip insertion when the name is empty.
            Console.WriteLine("Warning: Attempted to insert a form field with an empty name. Skipping insertion.");
            return;
        }

        // Name is valid; perform the insertion.
        insertAction();
    }
}
