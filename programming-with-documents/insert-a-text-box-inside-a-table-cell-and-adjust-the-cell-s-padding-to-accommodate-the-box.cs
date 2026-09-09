using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsTextBoxInTableCell
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and insert the first cell.
            builder.StartTable();
            builder.InsertCell();

            // Increase the cell padding so the text box does not touch the cell borders.
            // Padding values are in points.
            builder.CellFormat.SetPaddings(10, 10, 10, 10);

            // Insert a text box shape into the current cell.
            // Width and height are also specified in points.
            Shape shape = builder.InsertShape(ShapeType.TextBox, 150, 80);

            // Move the builder's cursor inside the text box and write the desired text.
            builder.MoveTo(shape.LastParagraph);
            builder.Write("This is a text box inside a table cell.");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the local file system.
            doc.Save("TextBoxInTableCell.docx");
        }
    }
}
