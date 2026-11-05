using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 pixel PNG image (semi‑transparent) from a base64 string.
        // This image will act as the watermark inside the table cell.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        string imagePath = "watermark.png";
        File.WriteAllBytes(imagePath, imageBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a merged cell.
        builder.StartTable();

        // First row – two normal cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row – first cell will be merged across two columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First; // Start merge.
        builder.Write("Merged Cell (watermark will be placed here)");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue merge.
        builder.EndRow();

        builder.EndTable();

        // Retrieve the table we just created.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Move the cursor to the first paragraph of the merged cell.
        builder.MoveTo(table.Rows[1].Cells[0].FirstParagraph);

        // Insert the watermark image as an inline shape so it does not disrupt layout.
        Shape watermark = builder.InsertImage(imagePath);
        watermark.WrapType = WrapType.Inline;
        // Set a reasonable width; height will preserve aspect ratio.
        watermark.Width = 100;

        // Save the document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully: " + outputPath);
        }
    }
}
