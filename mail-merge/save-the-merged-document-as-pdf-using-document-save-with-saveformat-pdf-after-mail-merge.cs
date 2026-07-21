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

        // Insert merge fields into the document.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FullName", "<FullName>");
        builder.Write(",");
        builder.Writeln();
        builder.Write("Welcome to ");
        builder.InsertField("MERGEFIELD Company", "<Company>");
        builder.Writeln(".");

        // Prepare data for a single record mail merge.
        string[] fieldNames = { "FullName", "Company" };
        object[] fieldValues = { "John Doe", "Acme Corp" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Define the output PDF file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedDocument.pdf");

        // Save the merged document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
