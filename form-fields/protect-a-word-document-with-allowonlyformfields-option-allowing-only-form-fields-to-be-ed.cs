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

        // Use DocumentBuilder to add content and form fields.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Please fill the form below:");
        builder.Write("Name: ");
        // Insert a text input form field.
        builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter name", 0);
        builder.InsertBreak(BreakType.ParagraphBreak);
        builder.Write("Agree to terms: ");
        // Insert a checkbox form field.
        builder.InsertCheckBox("AgreeField", false, 50);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Save the protected document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ProtectedFormFields.docx");
        doc.Save(outputPath);
    }
}
