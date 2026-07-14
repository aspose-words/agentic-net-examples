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

        // Apply a custom top border: blue color, single line, 2 pt width.
        // Use the Table.SetBorder method because Table.Borders is not available.
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Blue, true);

        // Define output path (relative to the executable folder).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTopBorderTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation – ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
