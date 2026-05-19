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

        // Example 1: Attempt to insert a text input form field with an empty name.
        string emptyName = string.Empty;
        try
        {
            InsertTextInputSafe(builder, emptyName, "Enter your text here", 50);
        }
        catch (ArgumentException ex)
        {
            // Log a warning instead of throwing.
            Console.WriteLine($"Warning: {ex.Message}");
        }

        // Insert a valid text input form field.
        InsertTextInputSafe(builder, "UserName", "Enter your name", 50);

        // Example 2: Attempt to insert a checkbox form field with an empty name.
        try
        {
            InsertCheckBoxSafe(builder, emptyName, false, 20);
        }
        catch (ArgumentException ex)
        {
            Console.WriteLine($"Warning: {ex.Message}");
        }

        // Insert a valid checkbox form field.
        InsertCheckBoxSafe(builder, "AgreeTerms", false, 20);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FormFieldsExample.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Inserts a text input form field after validating the field name.
    private static void InsertTextInputSafe(DocumentBuilder builder, string name, string placeholder, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Form field name cannot be empty.");

        // Insert the text input form field using the Aspose.Words API.
        builder.InsertTextInput(name, TextFormFieldType.Regular, "", placeholder, maxLength);
    }

    // Inserts a checkbox form field after validating the field name.
    private static void InsertCheckBoxSafe(DocumentBuilder builder, string name, bool defaultChecked, int size)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Form field name cannot be empty.");

        // Insert the checkbox form field using the Aspose.Words API.
        builder.InsertCheckBox(name, defaultChecked, size);
    }
}
