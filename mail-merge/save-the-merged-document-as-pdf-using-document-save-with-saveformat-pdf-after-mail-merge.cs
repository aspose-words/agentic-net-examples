using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple template with merge fields.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Perform mail merge with a single record.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello! This message was created with Aspose.Words mail merge." };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
