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

        // Build a large table. No explicit layout disabling APIs are needed;
        // Aspose.Words automatically defers layout calculations until the document is saved or UpdatePageLayout is called.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add many data rows.
        for (int i = 0; i < 2000; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 2");
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Force a layout refresh now that the table is complete.
        doc.UpdatePageLayout();

        // Save the document.
        string outputPath = "LargeTable.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
