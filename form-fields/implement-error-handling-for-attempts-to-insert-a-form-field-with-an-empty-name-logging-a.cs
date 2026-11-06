using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Attempt to insert a form field with an empty name.
        string emptyName = "";
        InsertTextInputSafe(builder, emptyName, "Default", 0, "");

        // Insert a valid text input form field.
        string validName = "CustomerName";
        InsertTextInputSafe(builder, validName, "John Doe", 0, "");

        // Save the document.
        doc.Save("FormFieldsOutput.docx");

        // Read and display the form fields that were successfully added.
        foreach (FormField field in doc.Range.FormFields)
        {
            // Ensure the field is not null before accessing its properties.
            if (field != null)
            {
                Console.WriteLine($"Field Name: {field.Name}, Result: {field.Result}");
            }
        }

        Console.WriteLine("Processing completed.");
    }

    private static void InsertTextInputSafe(DocumentBuilder builder, string name, string defaultText, int maxLength, string format)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            Console.WriteLine("Warning: Attempted to insert a form field with an empty name. Operation skipped.");
            return;
        }

        // Insert the text input form field using Aspose.Words API.
        builder.InsertTextInput(name, TextFormFieldType.Regular, format, defaultText, maxLength);
        Console.WriteLine($"Inserted text input form field '{name}' with default value '{defaultText}'.");
    }
}
