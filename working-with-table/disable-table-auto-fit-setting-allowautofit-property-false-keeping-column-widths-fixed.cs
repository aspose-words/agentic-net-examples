using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // First column – set a fixed width of 100 points.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Fixed width column 1");

        // Second column – set a fixed width of 150 points.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Fixed width column 2");
        builder.EndRow();

        // Disable automatic resizing of cells (auto‑fit).
        table.AllowAutoFit = false;

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableFixedAutoFit.docx");
        doc.Save(outputPath);
    }
}
