using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the DOT template.
        Document doc = new Document("Template.dot");

        // Create a DocumentBuilder to navigate the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the first cell of the first table (row 0, column 0).
        // Adjust the indices if the target cell is different.
        builder.MoveToCell(0, 0, 0, 0);

        // Define watermark options (font, size, color, layout, opacity).
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = System.Drawing.Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        // Although Watermark applies to the whole document, it will be visible inside the cell.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
