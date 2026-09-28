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

        // Attempt to insert a text input form field with an empty name.
        try
        {
            // This will throw an ArgumentException because the name is empty.
            builder.InsertTextInput(string.Empty, TextFormFieldType.Regular, "", "Default", 0);
        }
        catch (ArgumentException ex)
        {
            // Log a warning instead of terminating the program.
            Console.WriteLine($"Warning: Attempted to insert a form field with an empty name. Details: {ex.Message}");
        }

        // Insert a valid text input form field.
        string fieldName = "CustomerName";
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "John Doe", 0);
        builder.Writeln(); // Add a line break after the field.

        // Verify that the field exists and output its default result.
        FormField field = doc.Range.FormFields[fieldName];
        if (field != null)
        {
            Console.WriteLine($"Form field '{fieldName}' inserted with default value: '{field.Result}'.");
        }
        else
        {
            Console.WriteLine($"Error: Form field '{fieldName}' was not found after insertion.");
        }

        // Save the document.
        string outputPath = "FormFields.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
