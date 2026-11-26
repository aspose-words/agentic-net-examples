using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file path
        string outputPath = "MergedDocument.docx";

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build template with merge fields
        builder.Writeln("Hello ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Writeln(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln("!");

        // Prepare simple data source
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute mail merge
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save merged document as DOCX
        doc.Save(outputPath, SaveFormat.Docx);

        // Validation: ensure file exists and is not empty
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged document was not created.");

        if (new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("The merged document is empty.");
    }
}
