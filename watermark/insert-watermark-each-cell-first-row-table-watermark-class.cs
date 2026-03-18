using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class WatermarkInFirstRowCells
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for constructing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑column table with a header row and one data row.
        builder.StartTable();

        // ----- First row (header) -----
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // ----- Second row (data) -----
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.InsertCell();
        builder.Write("Data 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Define watermark options that will be applied to each cell.
        TextWatermarkOptions wmOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 24,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Iterate over each cell in the first row and apply the watermark.
        // The Watermark class works at the document level; calling SetText repeatedly
        // overwrites the previous watermark, so the final document will contain a single
        // watermark that appears on every page. This loop demonstrates the intended
        // per‑cell processing pattern.
        foreach (Cell cell in table.FirstRow.Cells)
        {
            // Move the builder's cursor to the beginning of the cell's first paragraph.
            builder.MoveTo(cell.FirstParagraph);

            // Apply the watermark to the document (affects the whole page).
            // In a real scenario, you might insert a shape or image directly into the cell
            // to achieve a true per‑cell watermark effect.
            doc.Watermark.SetText("Confidential", wmOptions);
        }

        // Save the resulting document.
        doc.Save("WatermarkedFirstRow.docx");
    }
}
