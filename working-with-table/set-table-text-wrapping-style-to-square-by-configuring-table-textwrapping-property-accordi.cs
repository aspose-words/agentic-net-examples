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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Set the table's text wrapping to "Around" (square style).
        table.TextWrapping = TextWrapping.Around;
        // Optional: define distances from surrounding text.
        table.AbsoluteHorizontalDistance = 20;
        table.AbsoluteVerticalDistance = 10;

        // Prepare output folder and file path.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWrapSquare.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
