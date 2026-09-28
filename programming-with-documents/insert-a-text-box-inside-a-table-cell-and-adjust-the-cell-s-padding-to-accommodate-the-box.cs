using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsTextBoxInTable
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

            // Get reference to the current cell.
            Cell cell = (Cell)builder.CurrentParagraph.ParentNode;

            // Adjust cell padding to accommodate the text box.
            cell.CellFormat.TopPadding = 20;
            cell.CellFormat.BottomPadding = 20;
            cell.CellFormat.LeftPadding = 20;
            cell.CellFormat.RightPadding = 20;

            // Create a text box shape.
            Shape textBox = new Shape(doc, ShapeType.TextBox)
            {
                Width = 200,
                Height = 100,
                WrapType = WrapType.Inline
            };

            // Add some text inside the text box.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "This is a text box inside a table cell.");
            para.AppendChild(run);
            textBox.AppendChild(para);

            // Insert the text box into the cell.
            cell.FirstParagraph.AppendChild(textBox);

            // End the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
