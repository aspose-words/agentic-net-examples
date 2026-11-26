using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the original document.
        string originalPath = "Original.docx";
        doc.Save(originalPath);

        // Iterate through all styles and modify table style definitions.
        foreach (Style style in doc.Styles)
        {
            if (style.Type == StyleType.Table)
            {
                TableStyle tableStyle = (TableStyle)style;
                // Change shading background color.
                tableStyle.Shading.BackgroundPatternColor = Color.LightYellow;
                // Change border color and width.
                tableStyle.Borders.Color = Color.DarkGreen;
                tableStyle.Borders.LineWidth = 1.5;
            }
        }

        // Save the modified document.
        string modifiedPath = "Modified.docx";
        doc.Save(modifiedPath);

        // Verify that both files were created.
        if (!File.Exists(originalPath) || !File.Exists(modifiedPath))
        {
            throw new Exception("Failed to create output files.");
        }

        Console.WriteLine($"Documents saved: {originalPath}, {modifiedPath}");
    }
}
