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

        // Start a table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        ApplyGradientShading(builder, TextureIndex.TextureDiagonalDown, Color.LightBlue);

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");
        ApplyGradientShading(builder, TextureIndex.TextureDiagonalUp, Color.LightCoral);

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        ApplyGradientShading(builder, TextureIndex.TextureHorizontal, Color.LightGreen);

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 4");
        ApplyGradientShading(builder, TextureIndex.TextureVertical, Color.LightGoldenrodYellow);

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "GradientShadingTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }

    // Helper method to apply gradient shading to the current cell.
    private static void ApplyGradientShading(DocumentBuilder builder, TextureIndex texture, Color backgroundColor)
    {
        // The builder is positioned inside the cell after InsertCell().
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        if (cell == null)
            throw new InvalidOperationException("Current node is not a cell.");

        // Set the shading texture and background pattern color.
        cell.CellFormat.Shading.Texture = texture;
        cell.CellFormat.Shading.BackgroundPatternColor = backgroundColor;
        // Optionally set a foreground color for the texture.
        cell.CellFormat.Shading.ForegroundPatternColor = Color.White;
    }
}
