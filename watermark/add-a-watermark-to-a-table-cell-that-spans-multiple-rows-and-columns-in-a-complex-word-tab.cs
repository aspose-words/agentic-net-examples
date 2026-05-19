using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Output document path.
        const string outputPath = "WatermarkCell.docx";

        // Create a tiny red PNG image from a Base64 string (no System.Drawing usage).
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/6ZcAAAAASUVORK5CYII=");
        const string imagePath = "watermark.png";
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Build a 4x4 table where the top‑left cell spans
        // two rows and two columns.
        // -------------------------------------------------
        Table table = builder.StartTable();

        // ----- First row -----
        // Cell (0,0) – start of the merged region.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // Cell (0,1) – horizontally merged with (0,0).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write(string.Empty);

        // Cell (0,2)
        builder.InsertCell();
        builder.Write("R0C2");

        // Cell (0,3)
        builder.InsertCell();
        builder.Write("R0C3");
        builder.EndRow();

        // ----- Second row -----
        // Cell (1,0) – vertically merged with (0,0).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(string.Empty);

        // Cell (1,1) – merged both horizontally and vertically.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(string.Empty);

        // Cell (1,2)
        builder.InsertCell();
        builder.Write("R1C2");

        // Cell (1,3)
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // ----- Remaining rows (simple cells) -----
        for (int r = 2; r < 4; r++)
        {
            for (int c = 0; c < 4; c++)
            {
                builder.InsertCell();
                builder.Write($"R{r}C{c}");
            }
            builder.EndRow();
        }

        // Reset cell formatting to avoid affecting later content.
        builder.CellFormat.ClearFormatting();

        // End the table construction.
        builder.EndTable();

        // -------------------------------------------------
        // Insert an image watermark into the merged cell.
        // -------------------------------------------------
        // Retrieve the merged cell (first row, first column).
        Cell mergedCell = table.Rows[0].Cells[0];

        // Move the builder's cursor to the first paragraph of the merged cell.
        builder.MoveTo(mergedCell.FirstParagraph);

        // Insert the image as a floating shape.
        Shape watermarkShape = builder.InsertImage(imagePath);

        // Configure the shape to behave like a watermark.
        watermarkShape.WrapType = WrapType.None;          // No text wrapping.
        watermarkShape.BehindText = true;                // Appear behind the cell text.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

        // Approximate sizing to cover the merged area.
        // Use the parent row's height for vertical dimension.
        watermarkShape.Width = mergedCell.CellFormat.Width * 2;                     // Approximate merged columns width.
        watermarkShape.Height = mergedCell.ParentRow.RowFormat.Height * 2;         // Approximate merged rows height.

        // -------------------------------------------------
        // Save the document.
        // -------------------------------------------------
        doc.Save(outputPath);
    }
}
