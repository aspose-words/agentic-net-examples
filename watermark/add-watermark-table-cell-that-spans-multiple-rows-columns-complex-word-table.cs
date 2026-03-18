using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

class TableCellWatermarkExample
{
    static void Main()
    {
        // Load an existing document that contains the complex table.
        Document doc = new Document("Input.docx");

        // Locate the target table. Adjust the index if the document has multiple tables.
        Table table = doc.FirstSection.Body.Tables[0];

        // Identify the cell that spans multiple rows and columns.
        // For this example we assume it is the cell at row 1, column 1 (zero‑based indexes).
        // In a real scenario you may need to search for cells with CellFormat.HorizontalMerge / VerticalMerge flags.
        Cell targetCell = table.Rows[0].Cells[0];

        // Ensure the cell contains at least one paragraph (required for inserting a shape).
        targetCell.EnsureMinimum();

        // Move the builder cursor to the first paragraph of the target cell.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert a WordArt shape that will act as the watermark.
        Shape watermarkShape = new Shape(doc, ShapeType.TextPlainText)
        {
            // Set the watermark text.
            TextPath = { Text = "CONFIDENTIAL", FontFamily = "Arial", Bold = true },

            // Position the shape relative to the cell.
            RelativeHorizontalPosition = RelativeHorizontalPosition.Margin,
            RelativeVerticalPosition = RelativeVerticalPosition.Margin,
            // Center the shape inside the cell.
            Left = 0,
            Top = 0,
            Width = targetCell.CellFormat.Width > 0 ? targetCell.CellFormat.Width : 200,
            Height = 50,

            // Rotate the text to give a typical watermark appearance.
            Rotation = -45,

            // Make the shape semi‑transparent.
            FillColor = Color.LightGray,
            // Use StrokeColor instead of the non‑existent OutlineColor.
            StrokeColor = Color.LightGray,
            // Set the shape behind the text.
            WrapType = WrapType.None,
            BehindText = true
        };

        // Add the shape to the cell's first paragraph.
        targetCell.FirstParagraph.AppendChild(watermarkShape);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
