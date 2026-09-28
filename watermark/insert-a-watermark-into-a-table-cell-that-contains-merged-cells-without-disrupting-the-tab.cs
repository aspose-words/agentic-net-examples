using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample watermark image (red square) from a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8/5+hHgAHggJ/Pi6ZVwAAAABJRU5ErkJggg==";
        byte[] imageBytes = Convert.FromBase64String(base64Png);
        const string imagePath = "watermark.png";
        File.WriteAllBytes(imagePath, imageBytes);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a merged cell.
        builder.StartTable();

        // First row - merged across two columns.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.Write(string.Empty);
        builder.EndRow();

        // Second row - normal cells.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Locate the merged cell (first row, first cell).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell mergedCell = table.Rows[0].Cells[0];

        // Move the builder cursor to the merged cell and insert the image as an inline watermark.
        builder.MoveTo(mergedCell.FirstParagraph);
        builder.InsertImage(imagePath);

        // Save the document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
