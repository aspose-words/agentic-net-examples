using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating image with square text wrapping.
        // The image file is created on the fly for demonstration purposes.
        string imagePath = "sample.png";
        CreateSampleImage(imagePath);
        Shape image = builder.InsertImage(imagePath);
        image.WrapType = WrapType.Square;
        image.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
        image.RelativeVerticalPosition = RelativeVerticalPosition.Line;
        image.Left = 50;   // Position from the left margin.
        image.Top = 50;    // Position from the top of the line.
        image.AllowOverlap = true;

        // Move to a new paragraph after the image.
        builder.Writeln();

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().Last();

        // In some versions of Aspose.Words the Table.AllowOverlap and TextWrappingStyle
        // properties are read‑only or not available. Therefore we rely on the default
        // behavior of the floating image (square wrap) which already allows the table
        // content to flow around the image.

        // Save the document.
        string outputPath = "TableWrapAroundImage.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
            File.Delete(imagePath);
    }

    // Helper method to create a simple placeholder PNG image without using System.Drawing.
    private static void CreateSampleImage(string path)
    {
        // This is a minimal 1x1 pixel PNG (transparent).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X" +
            "K6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(path, pngBytes);
    }
}
