using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class MailMergeExample
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

        // Prepare a data source for the mail merge.
        DataTable table = new DataTable("MailMergeData");
        table.Columns.Add("FirstName");
        table.Columns.Add("LastName");
        table.Columns.Add("Message");
        table.Rows.Add("John", "Doe", "Hello! This message was created with Aspose.Words mail merge.");

        // Execute the mail merge using the DataTable.
        doc.MailMerge.Execute(table);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MergedDocument.docx");

        // Save the merged document as DOCX.
        doc.Save(outputPath, SaveFormat.Docx);
    }
}
