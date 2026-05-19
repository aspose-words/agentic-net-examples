using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Define watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = System.Drawing.Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply a text watermark to the document.
        // The watermark will be visible on every page, including inside the table cells.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the resulting document.
        const string outputPath = "WatermarkedTableCell.docx";
        doc.Save(outputPath);
    }
}
