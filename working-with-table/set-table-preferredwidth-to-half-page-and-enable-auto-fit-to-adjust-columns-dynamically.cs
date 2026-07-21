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

        // Start building a table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row with varying content length.
        builder.InsertCell();
        builder.Write("This is a longer piece of text that will cause the column to expand if auto‑fit is disabled.");
        builder.InsertCell();
        builder.Write("Short");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set the table's preferred width to 50 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(50);

        // Ensure auto‑fit is enabled so column widths adjust dynamically based on content.
        table.AllowAutoFit = true; // Default is true, set explicitly for clarity.

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidthAutoFit.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved successfully.");
    }
}
