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

        // Build a simple 3x3 table.
        Table table = builder.StartTable();

        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Remove all existing borders (both outer and inner).
        table.ClearBorders();

        // Apply a custom outer border (single blue line, 2 points thick).
        // The last parameter 'true' overrides any existing cell borders.
        table.SetBorder(BorderType.Left,   LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Right,  LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Top,    LineStyle.Single, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Blue, true);

        // Save the document.
        string outputPath = "TableOuterBorder.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
