using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Saving;
using System.Drawing;

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
            Table table = builder.StartTable();
            builder.InsertCell();

            // Adjust the cell padding so the text box does not touch the borders.
            builder.CellFormat.SetPaddings(10, 10, 10, 10);
            // Optionally set a fixed width for the cell.
            builder.CellFormat.Width = 250;

            // Insert a text box shape into the current cell.
            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 80);
            // Add a paragraph with some text inside the text box.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "This is a text box inside a table cell.");
            para.AppendChild(run);
            textBox.AppendChild(para);

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document to the local file system.
            string outputPath = "TextBoxInTableCell.docx";
            doc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
