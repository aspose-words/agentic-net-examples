using System;
using System.Drawing;
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
        builder.EndRow();

        builder.EndTable();

        // Remove any existing borders (including inner cell borders).
        table.ClearBorders();

        // Apply a custom outer border style.
        // The 'true' flag overrides any explicit cell borders, ensuring inner borders stay hidden.
        table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Blue, true);

        // Save the document.
        string outputPath = "CustomOuterBorder.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
