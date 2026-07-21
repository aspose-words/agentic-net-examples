using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Apply a thick top border (3 points) and a thin bottom border (0.5 points).
        table.SetBorder(BorderType.Top, LineStyle.Single, 3.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 0.5, Color.Black, true);

        // Save the document to the local file system.
        string outputPath = "TableBorders.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
