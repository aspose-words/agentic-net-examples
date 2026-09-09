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

        // Add some introductory text.
        builder.Writeln("Please fill in the form field below:");

        // Insert a text input form field.
        // Parameters: name, type, format, default text, max length (0 = unlimited).
        builder.InsertTextInput("UserInput", TextFormFieldType.Regular, "", "Enter your text here", 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        const string outputPath = "ProtectedFormFields.docx";
        doc.Save(outputPath);

        // Inform that the file has been created (no user interaction required).
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
