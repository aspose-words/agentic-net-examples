using System;
using System.IO;
using Aspose.Words;

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
        builder.Writeln(".");

        // Perform mail merge with a single record.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
