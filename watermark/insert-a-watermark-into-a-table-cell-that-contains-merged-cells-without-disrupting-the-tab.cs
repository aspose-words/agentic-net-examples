using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a 3x3 table.
        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Merge the first row cells (0,0) to (0,2) into a single cell.
        Cell mergedCell = table.Rows[0].Cells[0];
        mergedCell.CellFormat.HorizontalMerge = CellMerge.First;
        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
        table.Rows[0].Cells[2].CellFormat.HorizontalMerge = CellMerge.Previous;

        // Replace any existing runs with custom text.
        mergedCell.FirstParagraph.Runs.Clear();
        mergedCell.FirstParagraph.AppendChild(new Run(doc, "Merged Cell Content"));

        // Create a simple PNG image to use as a watermark.
        string imagePath = "watermark.png";
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6V8AAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Insert the image watermark into the merged cell without breaking layout.
        Shape watermark = new Shape(doc, ShapeType.Image);
        watermark.ImageData.SetImage(imagePath);
        watermark.Width = 100;   // Adjust size as needed.
        watermark.Height = 30;
        watermark.WrapType = WrapType.None;
        watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermark.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        watermark.VerticalAlignment = VerticalAlignment.Center;
        watermark.HorizontalAlignment = HorizontalAlignment.Center;
        watermark.BehindText = true; // Appear as a watermark.

        // Add the watermark shape to the first paragraph of the merged cell.
        mergedCell.FirstParagraph.AppendChild(watermark);

        // Save the document.
        string outputPath = "TableCellWatermark.docx";
        doc.Save(outputPath);

        // Simple validation: ensure the output file exists.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully with cell watermark.");
        }
        else
        {
            Console.WriteLine("Failed to create the document.");
        }

        // Clean up the temporary image file.
        if (File.Exists(imagePath))
        {
            File.Delete(imagePath);
        }
    }
}
