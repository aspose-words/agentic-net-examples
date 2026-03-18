using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables; // <-- added namespace for Table class

class InsertPictureInTableCell
{
    static void Main()
    {
        // Path to the image file that will be inserted.
        const string imagePath = @"C:\Images\Sample.jpg";

        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for constructing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1‑row, 2‑cell table.
        builder.StartTable();
        builder.InsertCell();               // First cell.
        builder.InsertCell();               // Second cell.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Move the cursor to the first paragraph of the first cell.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating image shape into the cell.
        // Parameters: file name, horizontal position, left offset,
        // vertical position, top offset, width, height, wrap type.
        Shape picture = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.LeftMargin, 50,
            RelativeVerticalPosition.TopMargin, 100,
            100,   // width in points
            100,   // height in points
            WrapType.None);

        // Enable layout inside the cell so the shape moves with cell resizing.
        picture.IsLayoutInCell = true;

        // Save the resulting document.
        doc.Save(@"C:\Output\PictureInTableCell.docx");
    }
}
