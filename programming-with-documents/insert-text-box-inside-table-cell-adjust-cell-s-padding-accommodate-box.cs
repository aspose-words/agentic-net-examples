using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class TextBoxInTableCellExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert the first cell.
        builder.StartTable();
        builder.InsertCell();

        // Adjust the cell padding so the textbox fits comfortably.
        // Left, Top, Right, Bottom padding values are in points.
        builder.CellFormat.SetPaddings(10, 10, 10, 10);

        // Insert a textbox shape into the current cell.
        Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 80);
        // Ensure the shape is treated as part of the cell layout (default is true).
        textBoxShape.IsLayoutInCell = true;

        // Optionally configure the internal margins of the textbox.
        TextBox textBox = textBoxShape.TextBox;
        textBox.InternalMarginTop = 5;
        textBox.InternalMarginBottom = 5;
        textBox.InternalMarginLeft = 5;
        textBox.InternalMarginRight = 5;

        // Write some text inside the textbox.
        builder.MoveTo(textBoxShape.LastParagraph);
        builder.Write("Hello from the textbox!");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TextBoxInTableCell.docx");
    }
}
