using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Output file path
        string outputPath = "ShadedCell.docx";

        // Create a new document and a builder for it
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table
        Table table = builder.StartTable();

        // First cell - apply shading
        builder.InsertCell();
        // Access the first cell that was just created
        Cell firstCell = table.Rows[0].Cells[0];
        // Set a solid texture and a light gray background color
        firstCell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
        firstCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Write("Shaded cell");

        // Second cell - normal formatting
        builder.InsertCell();
        builder.Write("Normal cell");

        // End the row and the table
        builder.EndRow();
        builder.EndTable();

        // Save the document
        doc.Save(outputPath);

        // Verify that the file was created
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
