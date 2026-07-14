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

        // Insert merge fields that will be filled during the mail merge.
        builder.InsertField("MERGEFIELD FullName", "<FullName>");
        builder.Writeln();
        builder.InsertField("MERGEFIELD Address", "<Address>");

        // Define the field names and corresponding values for a single record.
        string[] fieldNames = { "FullName", "Address" };
        object[] fieldValues = { "John Doe", "123 Main St, Anytown" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as DOCX.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "MergedDocument.docx");
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
