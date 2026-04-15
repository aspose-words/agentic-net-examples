using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and insert the first cell.
            Table table = builder.StartTable();
            builder.InsertCell();

            // Increase the cell padding so the text box does not touch the cell borders.
            // Parameters: left, top, right, bottom (in points).
            builder.CellFormat.SetPaddings(10, 10, 10, 10);

            // Insert a text box shape into the current cell.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 100);
            // Add a paragraph with some text inside the text box.
            textBox.AppendChild(new Paragraph(doc));
            textBox.FirstParagraph.AppendChild(new Run(doc, "This is a text box inside a table cell."));

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to disk.
            doc.Save("TextBoxInTableCell.docx");
        }
    }
}
