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

        // Insert merge fields into the document.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Define field names and corresponding values for a single record.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello! This message was created with Aspose.Words mail merge." };

        // Perform the mail merge using the Execute(string[], object[]) overload.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as DOCX in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedResult.docx");
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
