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

        // Build a simple 2‑row, 3‑column table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.InsertCell();
        builder.Write("Row 1, Cell 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set the table's preferred width to 50 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(50);

        // Enable auto‑fit so column widths adjust dynamically based on content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TablePreferredWidth.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
