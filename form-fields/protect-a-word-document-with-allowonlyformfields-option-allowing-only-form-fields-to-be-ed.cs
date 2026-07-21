using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and a form field.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill in the form field below:");
        // Insert a text input form field named "UserName".
        builder.InsertTextInput("UserName", TextFormFieldType.Regular, "", "Enter your name", 0);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document.
        const string outputPath = "ProtectedFormFields.docx";
        doc.Save(outputPath);

        // Inform that the file has been created.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
