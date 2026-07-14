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

        // Use DocumentBuilder to add a paragraph with a text input form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the form below:");
        builder.Write("Name: ");
        // Insert a text input form field named "NameField" with placeholder text.
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name here", 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Define the output file path (in the current working directory).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedForm.docx");

        // Save the protected document.
        doc.Save(outputPath);
    }
}
