using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert merge fields that will be populated by the mail merge.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Define the data for the mail merge.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello! This document was created with Aspose.Words mail merge." };

        // Execute the mail merge for a single record.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as DOCX using the explicit SaveFormat.
        const string outputPath = "MergedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
