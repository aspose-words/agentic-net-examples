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

        // Attempt to insert a form field with an empty name.
        try
        {
            // This call is expected to throw an ArgumentException because the name is empty.
            builder.InsertTextInput(
                string.Empty,                     // name
                TextFormFieldType.Regular,        // type
                string.Empty,                     // format
                "Default",                        // default text
                0);                               // max length (0 = unlimited)

            Console.WriteLine("Unexpectedly inserted a form field with an empty name.");
        }
        catch (ArgumentException ex)
        {
            Console.WriteLine($"Warning: Attempted to insert a form field with an empty name. {ex.Message}");
        }

        // Insert a valid text input form field.
        builder.InsertParagraph();
        builder.InsertTextInput(
            "CustomerName",                     // name
            TextFormFieldType.Regular,          // type
            string.Empty,                       // format
            "John Doe",                         // default text
            0);                                 // max length (0 = unlimited)

        // Verify that the field was added.
        FormField field = doc.Range.FormFields["CustomerName"];
        if (field != null)
        {
            Console.WriteLine($"Form field '{field.Name}' inserted with default value '{field.Result}'.");
        }
        else
        {
            Console.WriteLine("Error: Expected form field not found.");
        }

        // Save the document.
        doc.Save("FormFields.docx");
        Console.WriteLine("Document saved as FormFields.docx.");
    }
}
