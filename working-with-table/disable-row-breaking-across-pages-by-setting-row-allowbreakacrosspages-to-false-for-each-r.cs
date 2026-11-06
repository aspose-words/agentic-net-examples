using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with three rows and two cells per row.
        builder.StartTable();
        for (int i = 0; i < 3; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i + 1}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i + 1}, Cell 2");

            // End the current row.
            builder.EndRow();
        }
        builder.EndTable();

        // Disable breaking rows across pages.
        foreach (Row row in doc.GetChildNodes(NodeType.Row, true))
        {
            // The property is on RowFormat, not directly on Row.
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to disk.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The document was not saved correctly.");
        }
    }
}
