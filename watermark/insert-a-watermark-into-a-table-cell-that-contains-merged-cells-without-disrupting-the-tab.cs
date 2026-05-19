using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell in the first row.
        builder.StartTable();

        // First cell – start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell with watermark");

        // Second cell – merged with the previous one.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.EndRow();

        // Second row – two normal cells.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // -----------------------------------------------------------------
        // Create a simple PNG image to use as a watermark.
        // The image is a 1x1 red pixel encoded as a byte array.
        // -----------------------------------------------------------------
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lKXcAAAAAElFTkSuQmCC");
        string imagePath = Path.Combine(Environment.CurrentDirectory, "watermark.png");
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // Insert the image into the merged cell as a watermark.
        // The shape is set to appear behind the text and not affect layout.
        // -----------------------------------------------------------------
        Table table = doc.FirstSection.Body.Tables[0];
        Cell mergedCell = table.Rows[0].Cells[0]; // The first cell of the merged range.

        // Insert the image shape into the cell.
        Shape watermarkShape = new Shape(doc, ShapeType.Image);
        watermarkShape.ImageData.SetImage(imagePath);
        watermarkShape.WrapType = WrapType.None;          // No text wrapping.
        watermarkShape.BehindText = true;                // Appear behind cell text.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        watermarkShape.Width = 100;   // Adjust size as needed.
        watermarkShape.Height = 100;

        // Add the shape to the first paragraph of the cell.
        mergedCell.FirstParagraph.AppendChild(watermarkShape);

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "output.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully: " + outputPath);
        }
    }
}
