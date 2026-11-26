using System;
using System.IO;
using System.Drawing;
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

        // Build a 3x3 table.
        builder.StartTable();

        // First row.
        builder.InsertCell(); builder.Write("R1C1");
        builder.InsertCell(); builder.Write("R1C2");
        builder.InsertCell(); builder.Write("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell(); builder.Write("R2C1");
        builder.InsertCell(); builder.Write("R2C2");
        builder.InsertCell(); builder.Write("R2C3");
        builder.EndRow();

        // Third row.
        builder.InsertCell(); builder.Write("R3C1");
        builder.InsertCell(); builder.Write("R3C2");
        builder.InsertCell(); builder.Write("R3C3");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Merge cells to create a cell that spans two rows and two columns (top‑left corner).
        // Horizontal merge.
        table.Rows[0].Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
        // Vertical merge.
        table.Rows[0].Cells[0].CellFormat.VerticalMerge = CellMerge.First;
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;

        // Prepare a simple placeholder PNG image (1x1 transparent pixel) as watermark source.
        // Base64 for a 1x1 transparent PNG.
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Move the builder cursor to the merged cell.
        builder.MoveTo(table.Rows[0].Cells[0].FirstParagraph);

        // Insert the image into the cell; this returns the Shape representing the image.
        Shape watermarkShape = builder.InsertImage(imageBytes);

        // Configure the shape to behave like a watermark inside the cell.
        watermarkShape.WrapType = WrapType.None;                     // No text wrapping.
        watermarkShape.BehindText = true;                           // Appear behind cell content.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        watermarkShape.VerticalAlignment = VerticalAlignment.Center;
        watermarkShape.HorizontalAlignment = HorizontalAlignment.Center;
        watermarkShape.Rotation = -45;                               // Slight rotation for visual effect.
        watermarkShape.Width = 200;                                 // Desired width.
        watermarkShape.Height = 200;                                // Desired height.
        watermarkShape.FillColor = Color.LightGray;                 // Light gray fill to simulate watermark.

        // Save the document.
        string outputPath = "WatermarkedTableCell.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
