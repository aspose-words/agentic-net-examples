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
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare data for a single record.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello! This message was created with Aspose.Words mail merge." };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedResult.docx");

        // Save the merged document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
