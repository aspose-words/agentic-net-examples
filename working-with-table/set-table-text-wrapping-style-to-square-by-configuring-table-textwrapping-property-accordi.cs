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

        // Give the table a preferred width so that it can be positioned as a floating object.
        table.PreferredWidth = PreferredWidth.FromPoints(300);

        // Set the text wrapping style to "Around" (square‑like wrapping).
        table.TextWrapping = TextWrapping.Around;

        // Optional: set distances from surrounding text.
        table.AbsoluteHorizontalDistance = 20;
        table.AbsoluteVerticalDistance = 20;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWrapText.docx");
        doc.Save(outputPath);
    }
}
