using System;
using System.IO;
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
            Cell cell = builder.InsertCell();

            // Increase the cell padding so the text box does not touch the borders.
            // Padding values are in points.
            builder.CellFormat.SetPaddings(leftPadding: 10, topPadding: 10, rightPadding: 10, bottomPadding: 10);

            // Create a text box shape.
            Shape textBoxShape = new Shape(doc, ShapeType.TextBox)
            {
                Width = 150,
                Height = 50,
                // Make the shape inline so it stays inside the cell.
                WrapType = WrapType.Inline
            };

            // A text box must contain at least one paragraph before we can move the builder into it.
            textBoxShape.AppendChild(new Paragraph(doc));

            // Insert the text box into the current cell.
            builder.InsertNode(textBoxShape);

            // Move the builder inside the text box and add some text.
            builder.MoveTo(textBoxShape.FirstParagraph);
            builder.Write("Aspose.Words TextBox");

            // Optionally add more content after the text box.
            builder.Writeln();
            builder.Write("Additional cell content.");

            // Finish the row and the table.
            builder.EndRow();
            builder.EndTable();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "TextBoxInTableCell.docx");
            doc.Save(outputPath);
        }
    }
}
