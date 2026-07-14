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

        // Start a table.
        Table table = builder.StartTable();

        // Build a 2x2 table and set each cell's orientation to vertical Asian script
        // (top‑to‑bottom). This achieves the same effect as setting Table.TextDirection.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                // Set the cell text orientation to vertical for Asian scripts.
                builder.CellFormat.Orientation = TextOrientation.VerticalFarEast;
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableTextDirection.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
