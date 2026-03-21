using System;
using System.Data;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // Create a destination document (simulating a DOCX converted from PDF)
        // -----------------------------------------------------------------
        Document destination = new Document();
        var destBuilder = new DocumentBuilder(destination);
        destBuilder.Writeln("=== Destination Document (converted from PDF) ===");
        destBuilder.Writeln("This content comes from the original PDF-derived DOCX.");
        destBuilder.Writeln();

        // -----------------------------------------------------------------
        // Create a mail‑merge template document (source)
        // -----------------------------------------------------------------
        Document source = new Document();
        var srcBuilder = new DocumentBuilder(source);
        srcBuilder.Writeln("=== Mail‑Merge Template ===");
        srcBuilder.Writeln("Dear <<FirstName>> <<LastName>>,");
        srcBuilder.Writeln("<<Message>>");
        srcBuilder.Writeln("Best regards,");
        srcBuilder.Writeln("Your Company");
        srcBuilder.Writeln();

        // -------------------------------------------------
        // Prepare a simple data source for the mail merge.
        // -------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("FirstName");
        data.Columns.Add("LastName");
        data.Columns.Add("Message");
        data.Rows.Add("John", "Doe", "Hello from mail merge!");
        data.Rows.Add("Jane", "Smith", "Another message.");

        // Use the first row as a single‑record data source.
        string[] fieldNames = data.Columns.Cast<DataColumn>()
                                          .Select(col => col.ColumnName)
                                          .ToArray();
        object[] fieldValues = data.Rows[0].ItemArray;

        // Execute the mail merge on the source document.
        source.MailMerge.Execute(fieldNames, fieldValues);

        // -------------------------------------------------
        // Append the mail‑merged document to the destination,
        // preserving the destination's styles.
        // -------------------------------------------------
        destination.AppendDocument(source, ImportFormatMode.UseDestinationStyles);

        // Save the combined document.
        string outputPath = "CombinedResult.docx";
        destination.Save(outputPath);
        Console.WriteLine($"Combined document saved to '{outputPath}'.");
    }
}
